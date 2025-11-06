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

	QUnit.module('Information formulas');

	QUnit.test("Test: \"CELL\"", function (assert) {

		ws.getRange2("J2").setValue("1");
		ws.getRange2("J3").setValue("test");
		ws.getRange2("J4").setValue("test2");
		ws.getRange2("J5").setValue("07/12/2000");
		ws.getRange2("J6").setValue("");
		
		oParser = new parserFormula('CELL("address",J3)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "$J$3");

		oParser = new parserFormula('CELL("address",J3:O12)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "$J$3");

		oParser = new parserFormula('CELL("col",J3)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 10);

		oParser = new parserFormula('CELL("col",J3:O12)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 10);

		oParser = new parserFormula('CELL("row",J3)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula('CELL("row",J3:O12)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula('CELL("color",J3)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula('CELL("color",J3:O12)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula('CELL("contents",J3)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "test");

		oParser = new parserFormula('CELL("contents",J3:O12)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "test");

		oParser = new parserFormula('CELL("contents",J4:O12)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "test2");

		oParser = new parserFormula('CELL("contents",J5:O12)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 36719);

		oParser = new parserFormula('CELL("prefix",J3)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "'");

		/*oParser = new parserFormula( 'CELL("prefix",J2)', "A1", ws );
		assert.ok( oParser.parse() );
		assert.strictEqual( oParser.calculate().getValue(), "" );*/

		oParser = new parserFormula('CELL("prefix",J6:O12)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");

		// address
		oParser = new parserFormula('CELL("address",J2)', "A1", ws);
		assert.ok(oParser.parse(), "Addres. Number type in cell.");
		assert.strictEqual(oParser.calculate().getValue(), "$J$2", "Addres. Number type in cell.");

		oParser = new parserFormula('CELL("address",J3)', "A1", ws);
		assert.ok(oParser.parse(), "Addres. String type in cell.");
		assert.strictEqual(oParser.calculate().getValue(), "$J$3", "Addres. String type in cell.");

		oParser = new parserFormula('CELL("address",J2:J3)', "A1", ws);
		assert.ok(oParser.parse(), "Addres. Cells range.");
		assert.strictEqual(oParser.calculate().getValue(), "$J$2", "Addres. Cells range.");

		oParser = new parserFormula('CELL("address",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(col,H23).");
		assert.strictEqual(oParser.calculate().getValue(), "$H$23", "Addres. Result of CELL(address,H23).");
		
		oParser = new parserFormula('CELL("address",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(address,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Addres. Result of CELL(address,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("address",12)', "A1", ws);
		assert.ok(oParser.parse(), "Addres. Cells range.");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Addres. Cells range.");

		oParser = new parserFormula('CELL("address",)', "A1", ws);
		assert.ok(oParser.parse(), "Addres. Cells range.");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Addres. Cells range.");

		oParser = new parserFormula('CELL("address",J)', "A1", ws);
		assert.ok(oParser.parse(), "Addres. Cells range.");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Addres. Cells range.");

		oParser = new parserFormula('CELL("address","J2")', "A1", ws);
		assert.ok(oParser.parse(), "Addres. Cells range.");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Addres. Cells range.");

		// col
		oParser = new parserFormula('CELL("col",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(col,J2).");
		assert.strictEqual(oParser.calculate().getValue(), 10, "Col. Result of CELL(col,J2).");

		oParser = new parserFormula('CELL("col",J2:J4)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(col,J2:J4).");
		assert.strictEqual(oParser.calculate().getValue(), 10, "Col. Result of CELL(col,J2:J4).");

		oParser = new parserFormula('CELL("col",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(col,H23).");
		assert.strictEqual(oParser.calculate().getValue(), 8, "Col. Result of CELL(col,H23).");
		
		oParser = new parserFormula('CELL("col",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(col,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Col. Result of CELL(col,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("col",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(col,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Col. Result of CELL(col,).");

		oParser = new parserFormula('CELL("col",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(col,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Col. Result of CELL(col,J).");

		oParser = new parserFormula('CELL("col","J2")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(col,'J2').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Result of CELL(col,'J2').");

		// color
		oParser = new parserFormula('CELL("color",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(color,J2).");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Color. Result of CELL(color,J2).");

		oParser = new parserFormula('CELL("color",J2:J4)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(color,J2:J4).");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Color. Result of CELL(color,J2:J4).");

		oParser = new parserFormula('CELL("color",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(color,H23).");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Color. Result of CELL(color,H23).");
		
		oParser = new parserFormula('CELL("color",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(color,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Color. Result of CELL(color,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("color",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(color,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Color. Result of CELL(color,).");

		oParser = new parserFormula('CELL("color",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(color,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Color. Result of CELL(color,J).");

		oParser = new parserFormula('CELL("color","J2")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(color,'J2').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Result of CELL(color,'J2').");

		// contents
		oParser = new parserFormula('CELL("contents",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(contents,J2).");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Contents. Result of CELL(contents,J2).");

		oParser = new parserFormula('CELL("contents",J2:J4)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(contents,J2:J4).");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Contents. Result of CELL(contents,J2:J4).");

		oParser = new parserFormula('CELL("contents",J5)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(contents,07/12/2000)");
		assert.strictEqual(oParser.calculate().getValue(), 36719, "Contents. Result of CELL(contents,07/12/2000).");

		oParser = new parserFormula('CELL("contents",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(contents,H23).");
		assert.strictEqual(oParser.calculate().getValue(), "", "Contents. Result of CELL(contents,H23).");
		
		oParser = new parserFormula('CELL("contents",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(contents,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Contents. Result of CELL(contents,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("contents",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(contents,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Contents. Result of CELL(contents,).");

		oParser = new parserFormula('CELL("contents",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(contents,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Contents. Result of CELL(contents,J).");

		oParser = new parserFormula('CELL("contents","J2")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(contents,'J2').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Contents. Result of CELL(contents,'J2').");

		// filename
		let sheetName = ws.sName;
		oParser = new parserFormula('CELL("filename",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(filename,J2).");
		assert.strictEqual(oParser.calculate().getValue(), "[TeSt.xlsx]" + sheetName, "filename. Result of CELL(filename,J2).");

		oParser = new parserFormula('CELL("filename",J2:J4)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(filename,J2:J4).");
		assert.strictEqual(oParser.calculate().getValue(), "[TeSt.xlsx]" + sheetName, "filename. Result of CELL(filename,J2:J4).");

		oParser = new parserFormula('CELL("filename",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(filename,H23).");
		assert.strictEqual(oParser.calculate().getValue(), "[TeSt.xlsx]" + sheetName, "filename. Result of CELL(filename,H23).");
		
		oParser = new parserFormula('CELL("filename",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(filename,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "filename. Result of CELL(filename,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("filename",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(filename,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "filename. Result of CELL(filename,).");

		oParser = new parserFormula('CELL("filename",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(filename,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "filename. Result of CELL(filename,J).");

		oParser = new parserFormula('CELL("filename","J2")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(filename,'J2').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "filename. Result of CELL(filename,'J2').");

		// format
		// G
		ws.getRange2("H2").setValue("50");
		// F0
		ws.getRange2("H3").setValue("0");
		ws.getRange2("H3").setNumFormat("0");
		// ,0
		ws.getRange2("H4").setValue("0");
		ws.getRange2("H4").setNumFormat("#,##0");
		// ,2
		ws.getRange2("H54").setValue("0.00");
		ws.getRange2("H54").setNumFormat("#,##0.00");
		// F2
		ws.getRange2("H5").setValue("0.00");
		ws.getRange2("H5").setNumFormat("0.00");
		// C0
		ws.getRange2("H66").setValue("0");
		ws.getRange2("H66").setNumFormat('#,##0;\\-#,##0');
		// C0 ms
		ws.getRange2("H6").setValue("0");
		ws.getRange2("H6").setNumFormat("$#,##0_);($#,##0)");
		// C0-
		ws.getRange2("H77").setValue("0");
		ws.getRange2("H77").setNumFormat('#,##0;[Red]\\-#,##0');
		// C0- ms 
		ws.getRange2("H7").setValue("0");
		ws.getRange2("H7").setNumFormat('$#,##0_);[Red]($#,##0)');
		// C2
		ws.getRange2("H88").setValue("0");
		ws.getRange2("H88").setNumFormat('#,##0.00;\-#,##0.00');
		// C2 ms
		ws.getRange2("H8").setValue("0");
		ws.getRange2("H8").setNumFormat('$#,##0.00_);($#,##0.00)');
		// C2-
		ws.getRange2("H99").setValue("0");
		ws.getRange2("H99").setNumFormat('#,##0.00;[Red]\-#,##0.00');
		// C2- ms
		ws.getRange2("H9").setValue("0");
		ws.getRange2("H9").setNumFormat('$#,##0.00_);[Red]($#,##0.00)');
		// P0
		ws.getRange2("H10").setValue("0");
		ws.getRange2("H10").setNumFormat("0%");
		// P2
		ws.getRange2("H11").setValue("0");
		ws.getRange2("H11").setNumFormat("0.00%");
		// S2
		ws.getRange2("H12").setValue("0");
		ws.getRange2("H12").setNumFormat("0.00E+00");
		// G
		ws.getRange2("H13").setValue("0");
		ws.getRange2("H13").setNumFormat("# ?/?");
		// G
		ws.getRange2("H113").setValue("0");
		ws.getRange2("H113").setNumFormat("# ??/??");
		// D1
		ws.getRange2("H14").setValue("10 Apr 20");
		ws.getRange2("H14").setNumFormat("dd/mm/yyyy");
		// D2
		ws.getRange2("H15").setValue("12-Jun");
		ws.getRange2("H15").setNumFormat("[$-9]d mmm;@");
		// D3
		ws.getRange2("H16").setValue("June-22");
		ws.getRange2("H16").setNumFormat("[$-9]mmm/yy;@");
		// D4
		ws.getRange2("H17").setValue("12/7/2022");
		ws.getRange2("H17").setNumFormat("m/d/yy;@");
		// D5
		ws.getRange2("H18").setValue("05/12");
		ws.getRange2("H18").setNumFormat("mm/dd;@");
		// D6
		ws.getRange2("H19").setValue("12:00:00 AM");
		ws.getRange2("H19").setNumFormat("h:mm:ss AM/PM");
		// D7
		ws.getRange2("H20").setValue("12:00 AM");
		ws.getRange2("H20").setNumFormat("h:mm AM/PM");
		// D8
		ws.getRange2("H21").setValue("12:00:00");
		ws.getRange2("H21").setNumFormat("h:mm:ss;@");
		// D9
		ws.getRange2("H22").setValue("12:00");
		ws.getRange2("H22").setNumFormat("h:mm;@");

		ws.getRange2("H23").setValue("{1,2,3,4,5}");
		

		oParser = new parserFormula('CELL("format",{0})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,{0}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "contents. Result of CELL(format,{0}).");

		oParser = new parserFormula('CELL("format",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "G", "contents. Result of CELL(format,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("format",{0;1;2;3})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H3:H22).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "contents. Result of CELL(format,{0;1;2;3}).");

		oParser = new parserFormula('CELL("format",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "contents. Result of CELL(format,).");

		oParser = new parserFormula('CELL("format",H3)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H3).");
		assert.strictEqual(oParser.calculate().getValue(), "F0", "contents. Result of CELL(format,0).");	// F0

		oParser = new parserFormula('CELL("format",H3:H22)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H3:H22).");
		assert.strictEqual(oParser.calculate().getValue(), "F0", "contents. Result of CELL(format,0).");	// F0

		oParser = new parserFormula('CELL("format",H4)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H4).");
		assert.strictEqual(oParser.calculate().getValue(), ",0", "contents. Result of CELL(format,H4).");	// ,0

		oParser = new parserFormula('CELL("format",H54)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H54).");
		assert.strictEqual(oParser.calculate().getValue(), ",2", "contents. Result of CELL(format,H54).");	// ,2

		oParser = new parserFormula('CELL("format",H5)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H5).");
		assert.strictEqual(oParser.calculate().getValue(), "F2", "contents. Result of CELL(format,H5).");	// F2

		oParser = new parserFormula('CELL("format",H6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H6).");
		assert.strictEqual(oParser.calculate().getValue(), "С0", "contents. Result of CELL(format,H6).");	// C0

		oParser = new parserFormula('CELL("format",H7)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H7).");
		assert.strictEqual(oParser.calculate().getValue(), "С0-", "contents. Result of CELL(format,H7).");	// C0-

		oParser = new parserFormula('CELL("format",H8)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H8).");
		assert.strictEqual(oParser.calculate().getValue(), "С2", "contents. Result of CELL(format,H8).");	// C2

		oParser = new parserFormula('CELL("format",H9)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H9).");
		assert.strictEqual(oParser.calculate().getValue(), "С2-", "contents. Result of CELL(format,H9).");	// C2-

		oParser = new parserFormula('CELL("format",H10)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H10).");
		assert.strictEqual(oParser.calculate().getValue(), "P0", "contents. Result of CELL(format,H10).");	// P0

		oParser = new parserFormula('CELL("format",H11)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H11).");
		assert.strictEqual(oParser.calculate().getValue(), "P2", "contents. Result of CELL(format,H11).");	// P2

		oParser = new parserFormula('CELL("format",H12)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H12).");
		assert.strictEqual(oParser.calculate().getValue(), "S2", "contents. Result of CELL(format,H12).");	// S2

		oParser = new parserFormula('CELL("format",H13)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H13).");
		assert.strictEqual(oParser.calculate().getValue(), "G", "contents. Result of CELL(format,H13).");	// "G"

		oParser = new parserFormula('CELL("format",H113)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,H113).");
		assert.strictEqual(oParser.calculate().getValue(), "G", "contents. Result of CELL(format,H113).");	// "G"

		oParser = new parserFormula('CELL("format",H2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,50).");
		assert.strictEqual(oParser.calculate().getValue(), "G", "Format. Result of CELL(format,50).");	// G

		oParser = new parserFormula('CELL("format",H14)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D1", "Format. Result of CELL(format,10 Apr 20).");	//D1

		oParser = new parserFormula('CELL("format",H15)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D2", "Format. Result of CELL(format,10 Apr 20).");	//D2

		oParser = new parserFormula('CELL("format",H16)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D3", "Format. Result of CELL(format,10 Apr 20).");	//D3

		oParser = new parserFormula('CELL("format",H17)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D4", "Format. Result of CELL(format,10 Apr 20).");	//D4

		oParser = new parserFormula('CELL("format",H18)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D5", "Format. Result of CELL(format,10 Apr 20).");	//D5

		oParser = new parserFormula('CELL("format",H19)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D6", "Format. Result of CELL(format,10 Apr 20).");	//D6
		
		oParser = new parserFormula('CELL("format",H20)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D7", "Format. Result of CELL(format,10 Apr 20).");	//D7

		oParser = new parserFormula('CELL("format",H21)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D8", "Format. Result of CELL(format,10 Apr 20).");	//D8

		oParser = new parserFormula('CELL("format",H22)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,10 Apr 20).");
		assert.strictEqual(oParser.calculate().getValue(), "D9", "Format. Result of CELL(format,10 Apr 20).");	//D9

		oParser = new parserFormula('CELL("format",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,1).");
		assert.strictEqual(oParser.calculate().getValue(), "G", "Format. Result of CELL(format,1).");

		oParser = new parserFormula('CELL("format",J3)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,string).");
		assert.strictEqual(oParser.calculate().getValue(), "G", "Format. Result of CELL(format,string).");

		oParser = new parserFormula('CELL("format",J5)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(format,07/12/2000).");
		assert.strictEqual(oParser.calculate().getValue(), "D4", "Format. Result of CELL(format,07/12/2000).");

		// parentheses
		oParser = new parserFormula('CELL("parentheses",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(parentheses,2).");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Parentheses. Result of CELL(parentheses,1).");

		oParser = new parserFormula('CELL("parentheses",J6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(parentheses,'').");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Parentheses. Result of CELL(parentheses,'').");

		oParser = new parserFormula('CELL("parentheses",J2:J6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(parentheses,J2:J6).");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Parentheses. Result of CELL(parentheses,J2:J6).");

		oParser = new parserFormula('CELL("parentheses",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(parentheses,H23).");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Parentheses. Result of CELL(parentheses,H23).");
		
		oParser = new parserFormula('CELL("parentheses",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(parentheses,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Parentheses. Result of CELL(parentheses,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("parentheses",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(parentheses,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Parentheses. Result of CELL(parentheses,).");
		
		oParser = new parserFormula('CELL("parentheses",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(parentheses,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Parentheses. Result of CELL(parentheses,J).")

		oParser = new parserFormula('CELL("parentheses","J")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(parentheses,'J').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Parentheses. Result of CELL(parentheses,'J').")

		// prefix
		oParser = new parserFormula('CELL("prefix",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(prefix,2).");
		assert.strictEqual(oParser.calculate().getValue(), "'", "Prefix. Result of CELL(prefix,1).");

		oParser = new parserFormula('CELL("prefix",J6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(prefix,'').");
		assert.strictEqual(oParser.calculate().getValue(), "", "Prefix. Result of CELL(prefix,'').");

		oParser = new parserFormula('CELL("prefix",J2:J6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(prefix,J2:J6).");
		assert.strictEqual(oParser.calculate().getValue(), "'", "Prefix. Result of CELL(prefix,J2:J6).");

		oParser = new parserFormula('CELL("prefix",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(prefix,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Prefix. Result of CELL(prefix,).");
		
		oParser = new parserFormula('CELL("prefix",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(prefix,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Prefix. Result of CELL(prefix,J).");

		oParser = new parserFormula('CELL("prefix",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(prefix,H23).");
		assert.strictEqual(oParser.calculate().getValue(), "'", "Prefix. Result of CELL(prefix,H23).");
		
		oParser = new parserFormula('CELL("prefix",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(prefix,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Prefix. Result of CELL(prefix,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("prefix","J")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(prefix,'J').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Prefix. Result of CELL(prefix,'J').");

		// protect
		oParser = new parserFormula('CELL("protect",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(protect,2).");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Protect. Result of CELL(protect,1).");

		oParser = new parserFormula('CELL("protect",J6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(protect,'').");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Protect. Result of CELL(protect,'').");

		oParser = new parserFormula('CELL("protect",J2:J6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(protect,J2:J6).");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Protect. Result of CELL(protect,J2:J6).");

		oParser = new parserFormula('CELL("protect",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(protect,H23).");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Protect. Result of CELL(protect,H23).");
		
		oParser = new parserFormula('CELL("protect",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(protect,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Protect. Result of CELL(protect,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("protect",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(protect,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Protect. Result of CELL(protect,).");
		
		oParser = new parserFormula('CELL("protect",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(protect,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Protect. Result of CELL(protect,J).")

		oParser = new parserFormula('CELL("protect","J")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(protect,'J').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Protect. Result of CELL(protect,'J').")

		// row
		oParser = new parserFormula('CELL("row",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(row,1).");
		assert.strictEqual(oParser.calculate().getValue(), 2, "Row. Result of CELL(row,1).");

		oParser = new parserFormula('CELL("row",J10)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(row,'').");
		assert.strictEqual(oParser.calculate().getValue(), 10, "Row. Result of CELL(row,'').");

		oParser = new parserFormula('CELL("row",B2:J5)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(row,B2:J5).");
		assert.strictEqual(oParser.calculate().getValue(), 2, "Row. Result of CELL(row,B2:J5).");

		oParser = new parserFormula('CELL("row",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(row,H23).");
		assert.strictEqual(oParser.calculate().getValue(), 23, "Row. Result of CELL(row,H23).");
		
		oParser = new parserFormula('CELL("row",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(row,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Row. Result of CELL(row,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("row",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(row,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Row. Result of CELL(row,).");

		oParser = new parserFormula('CELL("row",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(row,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Row. Result of CELL(row,J).")

		oParser = new parserFormula('CELL("row","J")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(row,'J').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Row. Result of CELL(row,'J').")

		// type
		oParser = new parserFormula('CELL("type",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(type,1).");
		assert.strictEqual(oParser.calculate().getValue(), "v", "Type. Result of CELL(type,1).");

		oParser = new parserFormula('CELL("type",J3)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(type,string).");
		assert.strictEqual(oParser.calculate().getValue(), "l", "Type. Result of CELL(type,string).");

		oParser = new parserFormula('CELL("type",J6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(type,'').");
		assert.strictEqual(oParser.calculate().getValue(), "b", "Type. Result of CELL(type,'').");

		oParser = new parserFormula('CELL("type",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(type,H23).");
		assert.strictEqual(oParser.calculate().getValue(), "l", "Type. Result of CELL(type,H23).");
		
		oParser = new parserFormula('CELL("type",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(type,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Type. Result of CELL(type,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("type",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(type,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Type. Result of CELL(type,).");

		oParser = new parserFormula('CELL("type",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(type,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Type. Result of CELL(type,J).");

		oParser = new parserFormula('CELL("type","J")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(type,'J').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Type. Result of CELL(type,'J').");

		// width
		oParser = new parserFormula('CELL("width",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(width,1).");
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 8, "Width. Result of CELL(width,1).");
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), "TRUE", "Width. Result of CELL(width,1).");

		oParser = new parserFormula('CELL("width",J3)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(width,string).");
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 8, "Width. Result of CELL(width,string).");
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), "TRUE", "Width. Result of CELL(width,string).");

		oParser = new parserFormula('CELL("width",J6)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(width,'').");
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 8, "Width. Result of CELL(width,'').");
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), "TRUE", "Width. Result of CELL(width,'').");

		oParser = new parserFormula('CELL("width",H23)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(width,H23).");
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 8, "Width. Result of CELL(width,H23).");
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), "TRUE", "Width. Result of CELL(width,H23).");
		
		oParser = new parserFormula('CELL("width",{1,2,3,4,5})', "A1", ws);
		assert.ok(oParser.parse(), "CELL(width,{1,2,3,4,5}).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Width. Result of CELL(width,{1,2,3,4,5}).");

		oParser = new parserFormula('CELL("width",)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(width,).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Width. Result of CELL(width,).");

		oParser = new parserFormula('CELL("width",J)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(width,J).");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Width. Result of CELL(width,J).");

		oParser = new parserFormula('CELL("width","J")', "A1", ws);
		assert.ok(oParser.parse(), "CELL(width,'J').");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "Width. Result of CELL(width,'J').");

		oParser = new parserFormula('CELL("fiLename",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(fiLename,J2).");
		assert.strictEqual(oParser.calculate().getValue(), "[TeSt.xlsx]" + sheetName, "fiLename. Result of CELL(filename,J2).");

		oParser = new parserFormula('CELL("FILENAME",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(FILENAME,J2).");
		assert.strictEqual(oParser.calculate().getValue(), "[TeSt.xlsx]" + sheetName, "FILENAME. Result of CELL(filename,J2).");

		oParser = new parserFormula('CELL("FILENAM",J2)', "A1", ws);
		assert.ok(oParser.parse(), "CELL(FILENAM,J2).");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "FILENAM. Result of CELL(filename,J2).");

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("12/12/2000");
		ws.getRange2("A102").setValue("12:00:00");
		ws.getRange2("A103").setValue("Test");
		ws.getRange2("A104").setValue("");
		ws.getRange2("A105").setValue("#N/A");
		ws.getRange2("A106").setValue("contents");
		ws.getRange2("A107").setValue("type");
        ws.getRange2("A108").setValue("#REF!");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Column1
		ws.getRange2("B601").setValue("contents"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("contents");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("1"); // TestName
		ws.getRange2("A202").setValue("contents"); // TestName1
		ws.getRange2("A206").setValue("1"); // TestNameArea
		ws.getRange2("A207").setValue("2"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("1") // TestName3D
		ws2.getRange2("A12").setValue("contents") // TestName3D1

		// Positive cases:
		// Case #1: String, Reference link. Testing basic functionality with valid parameters. Returns the address of cell reference.
		oParser = new parserFormula('CELL("address",A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("address",A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$A$100', 'Test: Positive case: String, Reference link. Testing basic functionality with valid parameters. Returns the address of cell reference.');
		// Case #2: String, Formula. Nested formula in reference parameter that resolves to valid reference.
		oParser = new parserFormula('CELL("contents",IF(1>0,A100,A101))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("contents",IF(1>0,A100,A101)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String, Formula. Nested formula in reference parameter that resolves to valid reference.');
		// Case #3: Formula, Reference link. Formula manipulating the info_type parameter but still valid.
		oParser = new parserFormula('CELL(LOWER("ADDRESS"),A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(LOWER("ADDRESS"),A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$A$100', 'Test: Positive case: Formula, Reference link. Formula manipulating the info_type parameter but still valid.');
		// Case #4: String, Date. Date value in cell - should return "v" for value type.
		oParser = new parserFormula('CELL("type",A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("type",A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'v', 'Test: Positive case: String, Date. Date value in cell - should return "v" for value type.');
		// Case #5: String, Time. Time value in cell - returns time as serial number.
		oParser = new parserFormula('CELL("contents",A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("contents",A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1) - 0, 0.5, 'Test: Positive case: String, Time. Time value in cell - returns time as serial number.');
		// Case #6: String, Boolean. Boolean value in cell - should return "l" for text.
		oParser = new parserFormula('CELL("type",A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("type",A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'l', 'Test: Positive case: String, Boolean. Boolean value in cell - should return "l" for text.');
		// Case #7: String, Empty. Empty cell - should return "b" for blank.
		oParser = new parserFormula('CELL("type",A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("type",A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'b', 'Test: Positive case: String, Empty. Empty cell - should return "b" for blank.');
		// Case #8: String, Error. Error value in cell - how CELL handles error values.
		oParser = new parserFormula('CELL("type",A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("type",A105) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'v', 'Test: Positive case: String, Error. Error value in cell - how CELL handles error values.');
		// Case #9: String, Table. Basic Table reference to test structured references.
		oParser = new parserFormula('CELL("contents",Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("contents",Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String, Table. Basic Table reference to test structured references.');
		// Case #10: String, Name. Named range reference - returns contents of first cell.
		oParser = new parserFormula('CELL("contents",TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("contents",TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String, Name. Named range reference - returns contents of first cell.');
		// Case #11: String, Name3D. 3D named range - tests address retrieval from 3D named range.
		oParser = new parserFormula('CELL("type",TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("type",TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'v', 'Test: Positive case: String, Name3D. 3D named range - tests address retrieval from 3D named range.');
		// Case #12: String, Ref3D. 3D reference to another sheet - tests content retrieval across sheets.
		oParser = new parserFormula('CELL("contents",Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("contents",Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String, Ref3D. 3D reference to another sheet - tests content retrieval across sheets.');
		// Case #13: String, Area3D. 3D range reference - tests address retrieval from multi-cell 3D range.
		oParser = new parserFormula('CELL("type",Sheet2!A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("type",Sheet2!A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'v', 'Test: Positive case: String, Area3D. 3D range reference - tests address retrieval from multi-cell 3D range.');
		// Case #14: String, Reference link. Case sensitivity test - "ADDRESS" instead of "address" - should still work but tests case sensitivity.
		oParser = new parserFormula('CELL("ADDRESS",A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("ADDRESS",A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$A$100', 'Test: Positive case: String, Reference link. Case sensitivity test - "ADDRESS" instead of "address" - should still work but tests case sensitivity.');
		// Case #15: String, Area. Multi-cell range with different data types - should return only first cell\'s contents.
		oParser = new parserFormula('CELL("contents",A100:A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("contents",A100:A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String, Area. Multi-cell range with different data types - should return only first cell\'s contents.');
		// Case #16: String, Name. Named range referring to multiple cells - tests handling of multi-cell names.
		oParser = new parserFormula('CELL("contents",TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("contents",TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String, Name. Named range referring to multiple cells - tests handling of multi-cell names.');
		// Case #17: String, Area. Multi-column range - tests handling of first cell\'s row only.
		oParser = new parserFormula('CELL("row",A100:B100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("row",A100:B100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: String, Area. Multi-column range - tests handling of first cell\'s row only.');
		// Case #18: Name, Reference link. Named range as info_type
		oParser = new parserFormula('CELL(TestName1,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(TestName1,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name, Reference link. Named range as info_type');
		// Case #19: Ref3D, Reference link. 3D reference as info_type
		oParser = new parserFormula('CELL(Sheet2!A2,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(Sheet2!A2,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D, Reference link. 3D reference as info_type');
		// Case #20: Area3D, Reference link. 3D range as info_type
		oParser = new parserFormula('CELL(Sheet2!A2:A2,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(Sheet2!A2:A2,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36872, 'Test: Positive case: Area3D, Reference link. 3D range as info_type');
		// Case #21: Table, Reference link. Table reference as info_type
		oParser = new parserFormula('CELL(Table1[Column2],A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(Table1[Column2],A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36872, 'Test: Positive case: Table, Reference link. Table reference as info_type');
		// Case #22: Reference link(2). Ref link as info_type
		oParser = new parserFormula('CELL(A106,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(A106,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link(2). Ref link as info_type');
		// Case #23: Area, Reference link. Area as info_type
		oParser = new parserFormula('CELL(A106:A107,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(A106:A107,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area, Reference link. Area as info_type');

		// Negative cases:
		// Case #1: Number, Reference link. info_type parameter as number instead of string - should return #VALUE!
		oParser = new parserFormula('CELL(1,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(1,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Reference link. info_type parameter as number instead of string - should return #VALUE!');
		// Case #2: Boolean, Reference link. info_type parameter as boolean instead of string - should return #VALUE!
		oParser = new parserFormula('CELL(TRUE,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(TRUE,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean, Reference link. info_type parameter as boolean instead of string - should return #VALUE!');
		// Case #4: String, Reference link. Completely invalid info_type parameter - should return #VALUE!
		oParser = new parserFormula('CELL("invalid_type",A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("invalid_type",A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Reference link. Completely invalid info_type parameter - should return #VALUE!');
		// Case #7: Empty, Reference link. Empty info_type parameter - should return #VALUE!
		oParser = new parserFormula('CELL(,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty, Reference link. Empty info_type parameter - should return #VALUE!');
		// Case #12: String, Reference link. Info_type with trailing space - tests whitespace sensitivity.
		oParser = new parserFormula('CELL("type ",A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("type ",A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Reference link. Info_type with trailing space - tests whitespace sensitivity.');
		// Case #13: String, Formula. Formula doesn't return #REF! error.
		oParser = new parserFormula('CELL("address",CHOOSE(2,A100,A108))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("address",CHOOSE(2,A100,A108)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$A$108', 'Test: Negative case: String, Formula. Formula doesnt return #REF! error.');
		// Case #15: Formula, Reference link. Formula returning invalid info_type - should return #VALUE!
		oParser = new parserFormula('CELL(IF(TRUE,"invalid","address"),A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(IF(TRUE,"invalid","address"),A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula, Reference link. Formula returning invalid info_type - should return #VALUE!');
		// Case #16: String, Reference link. Error value as reference parameter for "filename" - tests error handling.
		oParser = new parserFormula('CELL("filename",A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("filename",A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "[TeSt.xlsx]Sheet1", 'Test: Negative case: String, Reference link. Error value as reference parameter for "filename" - tests error handling.');
		// Case #19: String, Name. Non-existent named range - should return #NAME? error.
		oParser = new parserFormula('CELL("address",NonExistentName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("address",NonExistentName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: String, Name. Non-existent named range - should return #NAME? error.');
		// Case #23: String, Ref3D. Reference to non-existent sheet - should return #N/A error.
		oParser = new parserFormula('CELL("address",\'NonExistentSheet\'!A1)', 'A2', ws);
		assert.ok(oParser.parse() === false, 'Test: Formula CELL("address",\'NonExistentSheet\'!A1) is not parsed.');
		// assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String, Ref3D. Reference to non-existent sheet - should return #N/A error.');
		// Case #24: String, Ref3D. 3D reference with cell error - tests error propagation across sheets.
		oParser = new parserFormula('CELL("contents",Sheet2!#REF!)', 'A2', ws);
		assert.ok(oParser.parse() === false, 'Test: Formula CELL("contents",Sheet2!#REF!) is not parsed.');
		// assert.strictEqual(oParser.calculate().getValue(), '#REF!', 'Test: Negative case: String, Ref3D. 3D reference with cell error - tests error propagation across sheets.');
		// Case #25: String, Area3D. 3D range with non-existent sheet - should return #REF! error.
		oParser = new parserFormula('CELL("address",Sheet2!A1:\'NonExistentSheet\'!A1)', 'A2', ws);
		assert.ok(oParser.parse() === false, 'Test: Formula CELL("address",Sheet2!A1:NonExistentSheet!A1) is not parsed.');
		// assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Area3D. 3D range with non-existent sheet - should return #REF! error.');
		// Case #26: String, Area3D. 3D range with cell errors - tests error propagation in 3D ranges.
		oParser = new parserFormula('CELL("contents",Sheet2!#REF!:#REF!)', 'A2', ws);
		assert.ok(oParser.parse() === false, 'Test: Formula CELL("contents",Sheet2!#REF!:#REF!) is not parsed.');
		// assert.strictEqual(oParser.calculate().getValue(), '#REF!', 'Test: Negative case: String, Area3D. 3D range with cell errors - tests error propagation in 3D ranges.');
		// Case #29: String, Reference link. Reference to cell outside worksheet bounds - tests boundary handling.
		oParser = new parserFormula('CELL("address",A2000000)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("address",A2000000) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: String, Reference link. Reference to cell outside worksheet bounds - tests boundary handling.');
		// Case #32: Name, Reference link. Named range as info_type - should return #VALUE! if name contains non-text.
		oParser = new parserFormula('CELL(TestName,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(TestName,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name, Reference link. Named range as info_type - should return #VALUE! if name contains non-text.');
		// Case #33: Ref3D, Reference link. 3D reference as info_type - should return #VALUE! if cell contains non-text.
		oParser = new parserFormula('CELL(Sheet2!A1,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(Sheet2!A1,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, Reference link. 3D reference as info_type - should return #VALUE! if cell contains non-text.');
		// Case #34: Area3D, Reference link. 3D range as info_type - should return #VALUE! if range contains non-text.
		oParser = new parserFormula('CELL(Sheet2!A1:A1,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(Sheet2!A1:A1,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D, Reference link. 3D range as info_type - should return #VALUE! if range contains non-text.');
		// Case #35: Table, Reference link. Table reference as info_type - should return #VALUE! if column contains non-text.
		oParser = new parserFormula('CELL(Table1[Column1],A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL(Table1[Column1],A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table, Reference link. Table reference as info_type - should return #VALUE! if column contains non-text.');

		// Bounded cases:
		// Case #1: String, Formula. Using INDIRECT and ADDRESS to create reference to maximum row - tests boundary values.
		oParser = new parserFormula('CELL("row",INDIRECT(ADDRESS(1048576,1)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("row",INDIRECT(ADDRESS(1048576,1))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1048576, 'Test: Bounded case: String, Formula. Using INDIRECT and ADDRESS to create reference to maximum row - tests boundary values.');
		// Case #2: String, Formula. Using INDIRECT and ADDRESS to create reference to maximum column - tests boundary values.
		oParser = new parserFormula('CELL("col",INDIRECT(ADDRESS(1,16384)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CELL("col",INDIRECT(ADDRESS(1,16384))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16384, 'Test: Bounded case: String, Formula. Using INDIRECT and ADDRESS to create reference to maximum column - tests boundary values.');

        // Need to fix:
        // Case #8: String, Error. Error value in cell - how CELL handles error values.

	});

    QUnit.test("Test: \"ISBLANK\"", function (assert) {

		ws.getRange2("A202").setValue("");
		ws.getRange2("A203").setValue("test");

		oParser = new parserFormula('ISBLANK(A202)', "A1", ws);
		assert.ok(oParser.parse(), 'ISBLANK(A202)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISBLANK(A202)');

		oParser = new parserFormula('ISBLANK(A203)', "A1", ws);
		assert.ok(oParser.parse(), 'ISBLANK(A203)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISBLANK(A203)');

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Empty. Empty argument treated as empty cell, returns TRUE
		oParser = new parserFormula('ISBLANK(#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Empty. Empty argument treated as empty cell, returns TRUE');
		// Case #2: Number. Number input, returns FALSE
		oParser = new parserFormula('ISBLANK(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Number input, returns FALSE');
		// Case #3: String. String input, returns FALSE
		oParser = new parserFormula('ISBLANK("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String. String input, returns FALSE');
		// Case #4: Formula. Date formula returns serial number, returns FALSE
		oParser = new parserFormula('ISBLANK(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Date formula returns serial number, returns FALSE');
		// Case #5: Formula. Time formula returns serial number, returns FALSE
		oParser = new parserFormula('ISBLANK(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Time formula returns serial number, returns FALSE');
		// Case #6: Formula. Nested IF returning empty string, returns FALSE
		oParser = new parserFormula('ISBLANK(IF(TRUE,"",1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(IF(TRUE,"",1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested IF returning empty string, returns FALSE');
		// Case #7: Reference link. Reference to empty cell, returns TRUE
		oParser = new parserFormula('ISBLANK(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to empty cell, returns TRUE');
		// Case #8: Reference link. Reference to cell with number, returns FALSE
		oParser = new parserFormula('ISBLANK(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with number, returns FALSE');
		// Case #9: Area. Single-cell range with string, returns FALSE
		oParser = new parserFormula('ISBLANK(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell range with string, returns FALSE');
		// Case #10: Area. Multi-cell range, returns FALSE
		oParser = new parserFormula('ISBLANK(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Multi-cell range, returns FALSE');
		// Case #11: Array. Array with boolean, returns FALSE
		oParser = new parserFormula('ISBLANK({TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK({TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with boolean, returns FALSE');
		// Case #12: Array. Array with empty string, returns FALSE
		oParser = new parserFormula('ISBLANK({""})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK({""}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with empty string, returns FALSE');
		// Case #13: Name. Named range with empty cell, returns TRUE
		oParser = new parserFormula('ISBLANK(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with empty cell, returns TRUE');
		// Case #14: Name. Named range with number, returns FALSE
		oParser = new parserFormula('ISBLANK(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with number, returns FALSE');
		// Case #15: Name3D. 3D named range with empty cell, returns TRUE
		oParser = new parserFormula('ISBLANK(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with empty cell, returns TRUE');
		// Case #16: Name3D. 3D named range with string, returns FALSE
		oParser = new parserFormula('ISBLANK(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with string, returns FALSE');
		// Case #17: Ref3D. 3D reference to empty cell, returns TRUE
		oParser = new parserFormula('ISBLANK(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to empty cell, returns TRUE');
		// Case #18: Ref3D. 3D reference to cell with number, returns FALSE
		oParser = new parserFormula('ISBLANK(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with number, returns FALSE');
		// Case #19: Area3D. 3D single-cell range with string, returns FALSE
		oParser = new parserFormula('ISBLANK(Sheet2!A3:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Sheet2!A3:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with string, returns FALSE');
		// Case #20: Area3D. 3D multi-cell range, returns FALSE
		oParser = new parserFormula('ISBLANK(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D multi-cell range, returns FALSE');
		// Case #21: Table. Table column with empty cell, returns TRUE
		oParser = new parserFormula('ISBLANK(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with empty cell, returns TRUE');
		// Case #22: Table. Table column with number, returns FALSE
		oParser = new parserFormula('ISBLANK(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with number, returns FALSE');
		// Case #23: Formula. Nested SQRT formula, returns FALSE
		oParser = new parserFormula('ISBLANK(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested SQRT formula, returns FALSE');
		// Case #24: Formula. Volatile NOW formula, returns FALSE
		oParser = new parserFormula('ISBLANK(NOW())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(NOW()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Volatile NOW formula, returns FALSE');
		// Case #25: Formula. Nested SUM formula, returns FALSE
		oParser = new parserFormula('ISBLANK(SUM(0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(SUM(0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested SUM formula, returns FALSE');

		// Negative cases:
		// Case #1: Error. Error input propagates #N/A
		oParser = new parserFormula('ISBLANK(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Error. Error input propagates #N/A');
		// Case #2: Boolean. Boolean input, returns FALSE
		oParser = new parserFormula('ISBLANK(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean input, returns FALSE');
		// Case #3: Boolean. Boolean input, returns FALSE
		oParser = new parserFormula('ISBLANK(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean input, returns FALSE');
		// Case #4: String. String with space, returns FALSE
		oParser = new parserFormula('ISBLANK(" ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(" ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. String with space, returns FALSE');
		// Case #5: Formula. Empty string, returns FALSE
		oParser = new parserFormula('ISBLANK("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Empty string, returns FALSE');
		// Case #6: Reference link. Reference to cell with string space, returns FALSE
		oParser = new parserFormula('ISBLANK(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with string space, returns FALSE');
		// Case #7: Reference link. Reference to cell with error, returns FALSE
		oParser = new parserFormula('ISBLANK(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with error, returns FALSE');
		// Case #8: Area. Multi-cell range with mixed values, returns FALSE
		oParser = new parserFormula('ISBLANK(A107:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(A107:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range with mixed values, returns FALSE');
		// Case #9: Array. Multi-element array, returns FALSE
		oParser = new parserFormula('ISBLANK({FALSE,TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK({FALSE,TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Multi-element array, returns FALSE');
		// Case #10: Name. Named range with error, returns FALSE
		oParser = new parserFormula('ISBLANK(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with error, returns FALSE');
		// Case #11: Name. Named range with area, returns FALSE
		oParser = new parserFormula('ISBLANK(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with area, returns FALSE');
		// Case #12: Name3D. 3D named range with area, returns FALSE
		oParser = new parserFormula('ISBLANK(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with area, returns FALSE');
		// Case #13: Ref3D. 3D reference to cell with error, returns FALSE
		oParser = new parserFormula('ISBLANK(Sheet2!A6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Sheet2!A6) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Ref3D. 3D reference to cell with error, returns FALSE');
		// Case #14: Area3D. 3D multi-cell range with mixed values, returns FALSE
		oParser = new parserFormula('ISBLANK(Sheet2!A7:A8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Sheet2!A7:A8) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range with mixed values, returns FALSE');
		// Case #15: Table. Table column with string, returns FALSE
		oParser = new parserFormula('ISBLANK(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Table. Table column with string, returns FALSE');
		// Case #16: Formula. Formula resulting in #DIV/0!, returns FALSE
		oParser = new parserFormula('ISBLANK(DIV(1,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(DIV(1,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Formula resulting in #DIV/0!, returns FALSE');
		// Case #17: Formula. Nested IFERROR returning number, returns FALSE
		oParser = new parserFormula('ISBLANK(IFERROR(NA(),0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(IFERROR(NA(),0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested IFERROR returning number, returns FALSE');
		// Case #18: Number. Zero input, returns FALSE
		oParser = new parserFormula('ISBLANK(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Zero input, returns FALSE');
		// Case #19: String. String zero, returns FALSE
		oParser = new parserFormula('ISBLANK("0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK("0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. String zero, returns FALSE');
		// Case #20: Formula. Nested TEXT formula returning string, returns FALSE
		oParser = new parserFormula('ISBLANK(TEXT(DATE(2025,1,1),"yyyy"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(TEXT(DATE(2025,1,1),"yyyy")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested TEXT formula returning string, returns FALSE');

		// Bounded cases:
		// Case #1: Date. Minimum Excel date (serial number 1), returns FALSE
		oParser = new parserFormula('ISBLANK(DATE(1900,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(DATE(1900,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Date. Minimum Excel date (serial number 1), returns FALSE');
		// Case #2: Date. Maximum Excel date (serial number 2958465), returns FALSE
		oParser = new parserFormula('ISBLANK(DATE(9999,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(DATE(9999,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Date. Maximum Excel date (serial number 2958465), returns FALSE');
		// Case #3: Formula. Smallest positive Excel number, returns FALSE
		oParser = new parserFormula('ISBLANK(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Smallest positive Excel number, returns FALSE');
		// Case #4: Formula. Largest Excel number, returns FALSE
		oParser = new parserFormula('ISBLANK(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISBLANK(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Largest Excel number, returns FALSE');


		testArrayFormula2(assert, "ISBLANK", 1, 1);
	});

    QUnit.test("Test: \"ISERR\"", function (assert) {

		ws.getRange2("A202").setValue("");
		ws.getRange2("A203").setValue("#N/A");
		ws.getRange2("A204").setValue("#VALUE!");

		oParser = new parserFormula('ISERR(A202)', "A1", ws);
		assert.ok(oParser.parse(), 'ISERR(A202)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISERR(A202)');

		oParser = new parserFormula('ISERR(A203)', "A1", ws);
		assert.ok(oParser.parse(), 'ISERR(A203)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISERR(A203)');

		oParser = new parserFormula('ISERR(A203)', "A1", ws);
		assert.ok(oParser.parse(), 'ISERR(A203)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISERR(A203)');

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("#DIV/0!");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Error. #DIV/0! error, returns TRUE
		oParser = new parserFormula('ISERR(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #DIV/0! error, returns TRUE');
		// Case #2: Error. #VALUE! error, returns TRUE
		oParser = new parserFormula('ISERR(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #VALUE! error, returns TRUE');
		// Case #3: Error. #REF! error, returns TRUE
		oParser = new parserFormula('ISERR(#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #REF! error, returns TRUE');
		// Case #4: Error. #NAME? error, returns TRUE
		oParser = new parserFormula('ISERR(#NAME?)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(#NAME?) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #NAME? error, returns TRUE');
		// Case #5: Error. #NUM! error, returns TRUE
		oParser = new parserFormula('ISERR(#NUM!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(#NUM!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #NUM! error, returns TRUE');
		// Case #6: Error. #NULL! error, returns TRUE
		oParser = new parserFormula('ISERR(#NULL!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(#NULL!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #NULL! error, returns TRUE');
		// Case #7: Formula. Formula resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERR(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Formula resulting in #DIV/0!, returns TRUE');
		// Case #8: Formula. Nested formula resulting in #NUM!, returns TRUE
		oParser = new parserFormula('ISERR(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested formula resulting in #NUM!, returns TRUE');
		// Case #9: Reference link. Reference to cell with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERR(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with #DIV/0!, returns TRUE');
		// Case #10: Reference link. Reference to cell with #VALUE!, returns TRUE
		oParser = new parserFormula('ISERR(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with #VALUE!, returns TRUE');
		// Case #11: Area. Single-cell range with #REF!, returns TRUE
		oParser = new parserFormula('ISERR(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell range with #REF!, returns TRUE');
		// Case #12: Area. Multi-cell range with error, returns TRUE
		oParser = new parserFormula('ISERR(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Multi-cell range with error, returns TRUE');
		// Case #13: Array. Array with error value, returns TRUE
		oParser = new parserFormula('ISERR({#DIV/0!})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR({#DIV/0!}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with error value, returns TRUE');
		// Case #14: Name. Named range with #NUM!, returns TRUE
		oParser = new parserFormula('ISERR(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with #NUM!, returns TRUE');
		// Case #15: Name3D. 3D named range with #NULL!, returns TRUE
		oParser = new parserFormula('ISERR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with #NULL!, returns TRUE');
		// Case #16: Ref3D. 3D reference to cell with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERR(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with #DIV/0!, returns TRUE');
		// Case #17: Area3D. 3D single-cell range with #VALUE!, returns TRUE
		oParser = new parserFormula('ISERR(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with #VALUE!, returns TRUE');
		// Case #18: Table. Table column with #REF!, returns TRUE
		oParser = new parserFormula('ISERR(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with #REF!, returns TRUE');
		// Case #19: Formula. Nested IF resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERR(IF(TRUE,1/0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(IF(TRUE,1/0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF resulting in #DIV/0!, returns TRUE');
		// Case #20: Formula. Formula causing #VALUE!, returns TRUE
		oParser = new parserFormula('ISERR("text"+1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR("text"+1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Formula causing #VALUE!, returns TRUE');
		// Case #21: Reference link. Reference to cell with formula causing #NAME?, returns TRUE
		oParser = new parserFormula('ISERR(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with formula causing #NAME?, returns TRUE');
		// Case #22: Area3D. 3D multi-cell range with #NUM!, returns TRUE
		oParser = new parserFormula('ISERR(Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(Sheet2!A3:A4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D multi-cell range with #NUM!, returns TRUE');
		// Case #23: Formula. Nested LN formula causing #NUM!, returns TRUE
		oParser = new parserFormula('ISERR(LN(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(LN(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested LN formula causing #NUM!, returns TRUE');
		// Case #24: Name. Named range with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERR(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with #DIV/0!, returns TRUE');
		// Case #25: Table. Table column with formula causing #VALUE!, returns TRUE
		oParser = new parserFormula('ISERR(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with formula causing #VALUE!, returns TRUE');

		// Negative cases:
		// Case #1: Error. #N/A error, returns FALSE
		oParser = new parserFormula('ISERR(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Error. #N/A error, returns FALSE');
		// Case #2: Number. Number input, returns FALSE
		oParser = new parserFormula('ISERR(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Number input, returns FALSE');
		// Case #3: String. String input, returns FALSE
		oParser = new parserFormula('ISERR("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. String input, returns FALSE');
		// Case #4: Empty. Empty argument, returns FALSE
		oParser = new parserFormula('ISERR("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Empty. Empty argument, returns FALSE');
		// Case #5: Boolean. Boolean input, returns FALSE
		oParser = new parserFormula('ISERR(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean input, returns FALSE');
		// Case #6: Boolean. Boolean input, returns FALSE
		oParser = new parserFormula('ISERR(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean input, returns FALSE');
		// Case #7: Formula. Date formula returns serial number, returns FALSE
		oParser = new parserFormula('ISERR(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Date formula returns serial number, returns FALSE');
		// Case #8: Formula. Time formula returns serial number, returns FALSE
		oParser = new parserFormula('ISERR(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Time formula returns serial number, returns FALSE');
		// Case #9: Reference link. Reference to empty cell, returns FALSE
		oParser = new parserFormula('ISERR(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to empty cell, returns FALSE');
		// Case #10: Reference link. Reference to cell with number, returns FALSE
		oParser = new parserFormula('ISERR(A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with number, returns FALSE');
		// Case #11: Area. Single-cell range with string, returns FALSE
		oParser = new parserFormula('ISERR(A108:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A108:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Single-cell range with string, returns FALSE');
		// Case #12: Area. Multi-cell range with non-error values, returns FALSE
		oParser = new parserFormula('ISERR(A109:A110)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(A109:A110) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range with non-error values, returns FALSE');
		// Case #13: Array. Array with number, returns FALSE
		oParser = new parserFormula('ISERR({123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR({123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with number, returns FALSE');
		// Case #14: Name. Named range with empty cell, returns FALSE
		oParser = new parserFormula('ISERR(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with empty cell, returns FALSE');
		// Case #15: Name3D. 3D named range with number, returns FALSE
		oParser = new parserFormula('ISERR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with number, returns FALSE');
		// Case #16: Ref3D. 3D reference to cell with string, returns FALSE
		oParser = new parserFormula('ISERR(Sheet2!A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(Sheet2!A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Ref3D. 3D reference to cell with string, returns FALSE');
		// Case #17: Area3D. 3D multi-cell range with non-error values, returns FALSE
		oParser = new parserFormula('ISERR(Sheet2!A6:A7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(Sheet2!A6:A7) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range with non-error values, returns FALSE');
		// Case #18: Table. Table column with boolean, returns FALSE
		oParser = new parserFormula('ISERR(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Table. Table column with boolean, returns FALSE');
		// Case #19: Formula. Nested IF returning number, returns FALSE
		oParser = new parserFormula('ISERR(IF(TRUE,123,NA()))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(IF(TRUE,123,NA())) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested IF returning number, returns FALSE');
		// Case #20: Formula. Nested SQRT formula returning number, returns FALSE
		oParser = new parserFormula('ISERR(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested SQRT formula returning number, returns FALSE');

		// Bounded cases:
		// Case #1: Formula. Minimum Excel date divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERR(DATE(1900,1,1)/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(DATE(1900,1,1)/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Minimum Excel date divided by zero, returns #DIV/0!, returns TRUE');
		// Case #2: Formula. Maximum Excel date divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERR(DATE(9999,12,31)/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(DATE(9999,12,31)/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Maximum Excel date divided by zero, returns #DIV/0!, returns TRUE');
		// Case #3: Formula. Smallest positive Excel number divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERR(1E-307/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(1E-307/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Smallest positive Excel number divided by zero, returns #DIV/0!, returns TRUE');
		// Case #4: Formula. Largest Excel number causing overflow (#NUM!), returns TRUE
		oParser = new parserFormula('ISERR(9.99999999999999E+307*10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERR(9.99999999999999E+307*10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Largest Excel number causing overflow (#NUM!), returns TRUE');

		// Need to fix: area handle
		// Case #12: Area. Multi-cell range with error, returns TRUE
		// Case #22: Area3D. 3D multi-cell range with #NUM!, returns TRUE
		// Case #12: Area. Multi-cell range with non-error values, returns FALSE
		// Case #17: Area3D. 3D multi-cell range with non-error values, returns FALSE

		testArrayFormula2(assert, "ISERR", 1, 1);
	});

    QUnit.test("Test: \"ISERROR\"", function (assert) {

		ws.getRange2("A202").setValue("");
		ws.getRange2("A203").setValue("#N/A");

		oParser = new parserFormula('ISERROR(A202)', "A1", ws);
		assert.ok(oParser.parse(), 'ISERROR(A202)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISERROR(A202)');

		oParser = new parserFormula('ISERROR(A203)', "A1", ws);
		assert.ok(oParser.parse(), 'ISERROR(A203)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISERROR(A203)');

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("#DIV/0!");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Error. #DIV/0! error, returns TRUE
		oParser = new parserFormula('ISERROR(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #DIV/0! error, returns TRUE');
		// Case #2: Error. #VALUE! error, returns TRUE
		oParser = new parserFormula('ISERROR(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #VALUE! error, returns TRUE');
		// Case #3: Error. #REF! error, returns TRUE
		oParser = new parserFormula('ISERROR(#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #REF! error, returns TRUE');
		// Case #4: Error. #NAME? error, returns TRUE
		oParser = new parserFormula('ISERROR(#NAME?)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(#NAME?) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #NAME? error, returns TRUE');
		// Case #5: Error. #NUM! error, returns TRUE
		oParser = new parserFormula('ISERROR(#NUM!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(#NUM!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #NUM! error, returns TRUE');
		// Case #6: Error. #NULL! error, returns TRUE
		oParser = new parserFormula('ISERROR(#NULL!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(#NULL!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. #NULL! error, returns TRUE');
		// Case #7: Formula. Formula resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERROR(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Formula resulting in #DIV/0!, returns TRUE');
		// Case #8: Formula. Nested formula resulting in #NUM!, returns TRUE
		oParser = new parserFormula('ISERROR(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested formula resulting in #NUM!, returns TRUE');
		// Case #9: Reference link. Reference to cell with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERROR(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with #DIV/0!, returns TRUE');
		// Case #10: Reference link. Reference to cell with #VALUE!, returns TRUE
		oParser = new parserFormula('ISERROR(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with #VALUE!, returns TRUE');
		// Case #11: Area. Single-cell range with #REF!, returns TRUE
		oParser = new parserFormula('ISERROR(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell range with #REF!, returns TRUE');
		// Case #12: Area. Multi-cell range with error, returns TRUE
		oParser = new parserFormula('ISERROR(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Multi-cell range with error, returns TRUE');
		// Case #13: Array. Array with error value, returns TRUE
		oParser = new parserFormula('ISERROR({#DIV/0!})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR({#DIV/0!}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with error value, returns TRUE');
		// Case #14: Name. Named range with #NUM!, returns TRUE
		oParser = new parserFormula('ISERROR(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with #NUM!, returns TRUE');
		// Case #15: Name3D. 3D named range with #NULL!, returns TRUE
		oParser = new parserFormula('ISERROR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with #NULL!, returns TRUE');
		// Case #16: Ref3D. 3D reference to cell with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERROR(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with #DIV/0!, returns TRUE');
		// Case #17: Area3D. 3D single-cell range with #VALUE!, returns TRUE
		oParser = new parserFormula('ISERROR(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with #VALUE!, returns TRUE');
		// Case #18: Table. Table column with #REF!, returns TRUE
		oParser = new parserFormula('ISERROR(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with #REF!, returns TRUE');
		// Case #19: Formula. Nested IF resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERROR(IF(TRUE,1/0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(IF(TRUE,1/0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF resulting in #DIV/0!, returns TRUE');
		// Case #20: Formula. Formula causing #VALUE!, returns TRUE
		oParser = new parserFormula('ISERROR("text"+1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR("text"+1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Formula causing #VALUE!, returns TRUE');
		// Case #21: Reference link. Reference to cell with formula causing #NAME?, returns TRUE
		oParser = new parserFormula('ISERROR(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with formula causing #NAME?, returns TRUE');
		// Case #22: Area3D. 3D multi-cell range with #NUM!, returns TRUE
		oParser = new parserFormula('ISERROR(Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(Sheet2!A3:A4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D multi-cell range with #NUM!, returns TRUE');
		// Case #23: Formula. Nested LN formula causing #NUM!, returns TRUE
		oParser = new parserFormula('ISERROR(LN(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(LN(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested LN formula causing #NUM!, returns TRUE');
		// Case #24: Name. Named range with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERROR(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with #DIV/0!, returns TRUE');
		// Case #25: Table. Table column with formula causing #VALUE!, returns TRUE
		oParser = new parserFormula('ISERROR(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with formula causing #VALUE!, returns TRUE');

		// Negative cases:
		// Case #1: Error. #N/A error, returns FALSE
		oParser = new parserFormula('ISERROR(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Error. #N/A error, returns FALSE');
		// Case #2: Number. Number input, returns FALSE
		oParser = new parserFormula('ISERROR(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Number input, returns FALSE');
		// Case #3: String. String input, returns FALSE
		oParser = new parserFormula('ISERROR("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. String input, returns FALSE');
		// Case #4: Empty. Empty argument, returns FALSE
		oParser = new parserFormula('ISERROR("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Empty. Empty argument, returns FALSE');
		// Case #5: Boolean. Boolean input, returns FALSE
		oParser = new parserFormula('ISERROR(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean input, returns FALSE');
		// Case #6: Boolean. Boolean input, returns FALSE
		oParser = new parserFormula('ISERROR(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean input, returns FALSE');
		// Case #7: Formula. Date formula returns serial number, returns FALSE
		oParser = new parserFormula('ISERROR(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Date formula returns serial number, returns FALSE');
		// Case #8: Formula. Time formula returns serial number, returns FALSE
		oParser = new parserFormula('ISERROR(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Time formula returns serial number, returns FALSE');
		// Case #9: Reference link. Reference to empty cell, returns FALSE
		oParser = new parserFormula('ISERROR(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to empty cell, returns FALSE');
		// Case #10: Reference link. Reference to cell with number, returns FALSE
		oParser = new parserFormula('ISERROR(A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with number, returns FALSE');
		// Case #11: Area. Single-cell range with string, returns FALSE
		oParser = new parserFormula('ISERROR(A108:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A108:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Single-cell range with string, returns FALSE');
		// Case #12: Area. Multi-cell range with non-error values, returns FALSE
		oParser = new parserFormula('ISERROR(A109:A110)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(A109:A110) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range with non-error values, returns FALSE');
		// Case #13: Array. Array with number, returns FALSE
		oParser = new parserFormula('ISERROR({123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR({123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with number, returns FALSE');
		// Case #14: Name. Named range with empty cell, returns FALSE
		oParser = new parserFormula('ISERROR(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with empty cell, returns FALSE');
		// Case #15: Name3D. 3D named range with number, returns FALSE
		oParser = new parserFormula('ISERROR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with number, returns FALSE');
		// Case #16: Ref3D. 3D reference to cell with string, returns FALSE
		oParser = new parserFormula('ISERROR(Sheet2!A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(Sheet2!A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Ref3D. 3D reference to cell with string, returns FALSE');
		// Case #17: Area3D. 3D multi-cell range with non-error values, returns FALSE
		oParser = new parserFormula('ISERROR(Sheet2!A6:A7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(Sheet2!A6:A7) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range with non-error values, returns FALSE');
		// Case #18: Table. Table column with boolean, returns FALSE
		oParser = new parserFormula('ISERROR(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Table. Table column with boolean, returns FALSE');
		// Case #19: Formula. Nested IF returning number, returns FALSE
		oParser = new parserFormula('ISERROR(IF(TRUE,123,NA()))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(IF(TRUE,123,NA())) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested IF returning number, returns FALSE');
		// Case #20: Formula. Nested SQRT formula returning number, returns FALSE
		oParser = new parserFormula('ISERROR(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested SQRT formula returning number, returns FALSE');

		// Bounded cases:
		// Case #1: Formula. Minimum Excel date divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERROR(DATE(1900,1,1)/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(DATE(1900,1,1)/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Minimum Excel date divided by zero, returns #DIV/0!, returns TRUE');
		// Case #2: Formula. Maximum Excel date divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERROR(DATE(9999,12,31)/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(DATE(9999,12,31)/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Maximum Excel date divided by zero, returns #DIV/0!, returns TRUE');
		// Case #3: Formula. Smallest positive Excel number divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISERROR(1E-307/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(1E-307/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Smallest positive Excel number divided by zero, returns #DIV/0!, returns TRUE');
		// Case #4: Formula. Largest Excel number causing overflow (#NUM!), returns TRUE
		oParser = new parserFormula('ISERROR(9.99999999999999E+307*10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISERROR(9.99999999999999E+307*10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Largest Excel number causing overflow (#NUM!), returns TRUE');

		// Need to fix: area handle
		// Case #12: Area. Multi-cell range with error, returns TRUE
		// Case #22: Area3D. 3D multi-cell range with #NUM!, returns TRUE
		// Case #12: Area. Multi-cell range with non-error values, returns FALSE
		// Case #17: Area3D. 3D multi-cell range with non-error values, returns FALSE

		testArrayFormula2(assert, "ISERROR", 1, 1);
	});

    QUnit.test("Test: \"ISEVEN\"", function (assert) {

		oParser = new parserFormula('ISEVEN(-1)', "A1", ws);
		assert.ok(oParser.parse(), 'ISEVEN(-1)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISEVEN(-1)');

		oParser = new parserFormula('ISEVEN(2.5)', "A1", ws);
		assert.ok(oParser.parse(), 'ISEVEN(2.5)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISEVEN(2.5)');

		oParser = new parserFormula('ISEVEN(5)', "A1", ws);
		assert.ok(oParser.parse(), 'ISEVEN(5)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISEVEN(5)');

		oParser = new parserFormula('ISEVEN(0)', "A1", ws);
		assert.ok(oParser.parse(), 'ISEVEN(0)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISEVEN(0)');

		oParser = new parserFormula('ISEVEN(12/23/2011)', "A1", ws);
		assert.ok(oParser.parse(), 'ISEVEN(12/23/2011)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISEVEN(12/23/2011)');

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Number. Integer input (even), returns TRUE
		oParser = new parserFormula('ISEVEN(2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Integer input (even), returns TRUE');
		// Case #2: Number. Integer input (odd), returns FALSE
		oParser = new parserFormula('ISEVEN(3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Integer input (odd), returns FALSE');
		// Case #3: Number. Zero (even), returns TRUE
		oParser = new parserFormula('ISEVEN(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Zero (even), returns TRUE');
		// Case #4: Number. Float truncated to 2 (even), returns TRUE
		oParser = new parserFormula('ISEVEN(2.7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(2.7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Float truncated to 2 (even), returns TRUE');
		// Case #5: Number. Float truncated to 3 (odd), returns FALSE
		oParser = new parserFormula('ISEVEN(3.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(3.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Float truncated to 3 (odd), returns FALSE');
		// Case #6: Formula. Nested ROUND formula (2, even), returns TRUE
		oParser = new parserFormula('ISEVEN(ROUND(2.3,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(ROUND(2.3,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested ROUND formula (2, even), returns TRUE');
		// Case #7: Formula. Nested SQRT formula (3, odd), returns FALSE
		oParser = new parserFormula('ISEVEN(SQRT(9))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(SQRT(9)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested SQRT formula (3, odd), returns FALSE');
		// Case #8: Reference link. Reference to cell with even number, returns TRUE
		oParser = new parserFormula('ISEVEN(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with even number, returns TRUE');
		// Case #9: Reference link. Reference to cell with odd number, returns FALSE
		oParser = new parserFormula('ISEVEN(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with odd number, returns FALSE');
		// Case #10: Area. Single-cell range with even number, returns TRUE
		oParser = new parserFormula('ISEVEN(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Single-cell range with even number, returns TRUE');
		// Case #11: Area. Multi-cell range with number, returns TRUE (first cell truncated to 2)
		oParser = new parserFormula('ISEVEN(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Multi-cell range with number, returns TRUE (first cell truncated to 2)');
		// Case #12: Array. Array with single even number, returns TRUE
		oParser = new parserFormula('ISEVEN({4})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN({4}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with single even number, returns TRUE');
		// Case #13: Array. Array with multiple numbers, evaluates first (3, odd), returns FALSE
		oParser = new parserFormula('ISEVEN({3,4})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN({3,4}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with multiple numbers, evaluates first (3, odd), returns FALSE');
		// Case #14: Name. Named range with even number, returns TRUE
		oParser = new parserFormula('ISEVEN(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name. Named range with even number, returns TRUE');
		// Case #15: Name3D. 3D named range with odd number, returns FALSE
		oParser = new parserFormula('ISEVEN(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D. 3D named range with odd number, returns FALSE');
		// Case #16: Ref3D. 3D reference to cell with even number, returns TRUE
		oParser = new parserFormula('ISEVEN(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with even number, returns TRUE');
		// Case #17: Area3D. 3D single-cell range with odd number, returns FALSE
		oParser = new parserFormula('ISEVEN(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area3D. 3D single-cell range with odd number, returns FALSE');
		// Case #18: Table. Table column with even number, returns TRUE
		oParser = new parserFormula('ISEVEN(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with even number, returns TRUE');
		// Case #19: Formula. Nested IF returning even number, returns TRUE
		oParser = new parserFormula('ISEVEN(IF(TRUE,4,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(IF(TRUE,4,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF returning even number, returns TRUE');
		// Case #20: Date. Date (serial number 45654, even), returns TRUE
		oParser = new parserFormula('ISEVEN(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Date. Date (serial number 45654, even), returns TRUE');
		// Case #21: Time. Time scaled to integer (500000, even), returns TRUE
		oParser = new parserFormula('ISEVEN(TIME(12,0,0)*1000000)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(TIME(12,0,0)*1000000) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Time. Time scaled to integer (500000, even), returns TRUE');
		// Case #22: String. String convertible to even number, returns TRUE
		oParser = new parserFormula('ISEVEN("4")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN("4") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. String convertible to even number, returns TRUE');
		// Case #23: Formula. Volatile NOW formula (serial number, e.g., 45654), returns TRUE
		oParser = new parserFormula('ISEVEN(NOW())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(NOW()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Volatile NOW formula (serial number, e.g., 45654), returns TRUE');
		// Case #24: Name. Named range with float truncated to even, returns TRUE
		oParser = new parserFormula('ISEVEN(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name. Named range with float truncated to even, returns TRUE');
		// Case #25: Table. Table column with odd number, returns FALSE
		oParser = new parserFormula('ISEVEN(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Table. Table column with odd number, returns FALSE');

		// Negative cases:
		// Case #1: Error. Error input (#DIV/0!), returns #VALUE!
		oParser = new parserFormula('ISEVEN(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Error. Error input (#DIV/0!), returns #VALUE!');
		// Case #2: Error. Error input (#VALUE!), returns #VALUE!
		oParser = new parserFormula('ISEVEN(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error. Error input (#VALUE!), returns #VALUE!');
		// Case #3: Error. Error input (#N/A), returns #VALUE!
		oParser = new parserFormula('ISEVEN(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error input (#N/A), returns #VALUE!');
		// Case #4: Error. Error input (#REF!), returns #VALUE!
		oParser = new parserFormula('ISEVEN(#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#REF!', 'Test: Negative case: Error. Error input (#REF!), returns #VALUE!');
		// Case #5: String. Non-numeric string, returns #VALUE!
		oParser = new parserFormula('ISEVEN("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string, returns #VALUE!');
		// Case #6: Empty. Empty argument, returns #VALUE!
		oParser = new parserFormula('ISEVEN("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty argument, returns #VALUE!');
		// Case #7: Boolean. Boolean input, returns #VALUE!
		oParser = new parserFormula('ISEVEN(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean input, returns #VALUE!');
		// Case #8: Boolean. Boolean input, returns #VALUE!
		oParser = new parserFormula('ISEVEN(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(FALSE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean input, returns #VALUE!');
		// Case #9: Formula. Formula causing #VALUE!, returns #VALUE!
		oParser = new parserFormula('ISEVEN("text"+1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN("text"+1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Formula causing #VALUE!, returns #VALUE!');
		// Case #10: Reference link. Reference to cell with error, returns #VALUE!
		oParser = new parserFormula('ISEVEN(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with error, returns #VALUE!');
		// Case #11: Reference link. Reference to cell with non-numeric string, returns #VALUE!
		oParser = new parserFormula('ISEVEN(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with non-numeric string, returns #VALUE!');
		// Case #12: Area. Single-cell range with error, returns #VALUE!
		oParser = new parserFormula('ISEVEN(A107:A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(A107:A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Single-cell range with error, returns #VALUE!');
		// Case #13: Area. Multi-cell range with non-numeric, returns #VALUE!
		oParser = new parserFormula('ISEVEN(A108:A109)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(A108:A109) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range with non-numeric, returns #VALUE!');
		// Case #14: Array. Array with non-numeric string, returns #VALUE!
		oParser = new parserFormula('ISEVEN({"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN({"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array. Array with non-numeric string, returns #VALUE!');
		// Case #15: Name. Named range with error, returns #VALUE!
		oParser = new parserFormula('ISEVEN(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name. Named range with error, returns #VALUE!');
		// Case #16: Name3D. 3D named range with non-numeric string, returns #VALUE!
		oParser = new parserFormula('ISEVEN(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name3D. 3D named range with non-numeric string, returns #VALUE!');
		// Case #17: Ref3D. 3D reference to cell with error, returns #VALUE!
		oParser = new parserFormula('ISEVEN(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to cell with error, returns #VALUE!');
		// Case #18: Area3D. 3D multi-cell range with non-numeric, returns #VALUE!
		oParser = new parserFormula('ISEVEN(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D multi-cell range with non-numeric, returns #VALUE!');
		// Case #19: Table. Table column with error, returns #VALUE!
		oParser = new parserFormula('ISEVEN(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with error, returns #VALUE!');
		// Case #20: Formula. Nested formula causing #NUM!, returns #VALUE!
		oParser = new parserFormula('ISEVEN(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Nested formula causing #NUM!, returns #VALUE!');

		// Bounded cases:
		// Case #1: Number. Smallest positive integer (odd), returns FALSE
		oParser = new parserFormula('ISEVEN(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Smallest positive integer (odd), returns FALSE');
		// Case #2: Number. Smallest negative integer (odd), returns FALSE
		oParser = new parserFormula('ISEVEN(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Smallest negative integer (odd), returns FALSE');
		// Case #3: Number. Largest even 32-bit integer, returns TRUE
		oParser = new parserFormula('ISEVEN(2^31-2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(2^31-2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Largest even 32-bit integer, returns TRUE');
		// Case #4: Number. Largest negative odd 32-bit integer, returns FALSE
		oParser = new parserFormula('ISEVEN(-(2^31-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(-(2^31-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Largest negative odd 32-bit integer, returns FALSE');
		// Case #5: Formula. Largest Excel number (truncated to integer), returns FALSE
		oParser = new parserFormula('ISEVEN(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISEVEN(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Largest Excel number (truncated to integer), returns FALSE');

		// Need to fix: boolean handle
		// Case #7: Boolean. Boolean input, returns #VALUE!
		// Case #8: Boolean. Boolean input, returns #VALUE!

		testArrayFormula2(assert, "ISEVEN", 1, 1, true, null);
	});

    QUnit.test("Test: \"ISFORMULA\"", function (assert) {

		ws.getRange2("C150").setValue("=TODAY()");
		ws.getRange2("C151").setValue("7");
		ws.getRange2("C152").setValue("Hello, world!");
		ws.getRange2("C153").setValue("=3/0");

		oParser = new parserFormula("ISFORMULA(C150)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().toString(), "TRUE");

		oParser = new parserFormula("ISFORMULA(C151)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().toString(), "FALSE");

		oParser = new parserFormula("ISFORMULA(C152)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().toString(), "FALSE");

		oParser = new parserFormula("ISFORMULA(C153)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().toString(), "TRUE");

		testArrayFormulaEqualsValues(assert, "FALSE,FALSE,FALSE,#N/A;FALSE,FALSE,FALSE,#N/A;#N/A,#N/A,#N/A,#N/A", "ISFORMULA(A1:C2)");
		testArrayFormulaEqualsValues(assert, "FALSE,FALSE,#N/A,#N/A;FALSE,FALSE,#N/A,#N/A;FALSE,FALSE,#N/A,#N/A", "ISFORMULA(A1:B1)");
		testArrayFormulaEqualsValues(assert, "FALSE,FALSE,FALSE,FALSE;FALSE,FALSE,FALSE,FALSE;FALSE,FALSE,FALSE,FALSE", "ISFORMULA(A1)");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("SQRT");
		ws.getRange2("A101").setValue("SUM");
		ws.getRange2("A104").setValue("Text");
		// For area
		ws.getRange2("A102").setValue("=SUM(1)");
		ws.getRange2("A103").setValue('=TEXT("1s")');
		ws.getRange2("A105").setValue("NonFormula");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("SUM(1)"); // Text (Column1)
		ws.getRange2("B601").setValue("=SUM(1)"); // Formula (Column2)
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
		// Case #1: Reference link. Reference to cell with formula returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with formula returns TRUE. 1 of 1 argument used.');
		// Case #2: Reference link. Reference to cell with constant returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with constant returns FALSE. 1 of 1 argument used.');
		// Case #3: Formula. Formula returning reference to formula cell returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(INDIRECT("A100"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(INDIRECT("A100")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Formula returning reference to formula cell returns TRUE. 1 of 1 argument used.');
		// Case #4: Reference link. Reference to cell with text returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with text returns FALSE. 1 of 1 argument used.');
		// Case #5: Reference link. Reference to empty cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to empty cell returns FALSE. 1 of 1 argument used.');
		// Case #6: Area. Single-cell range with formula returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Single-cell range with formula returns TRUE. 1 of 1 argument used.');
		// Case #7: Name. Named range referencing formula cell returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range referencing formula cell returns TRUE. 1 of 1 argument used.');
		// Case #8: Name3D. 3D named range referencing formula cell returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range referencing formula cell returns TRUE. 1 of 1 argument used.');
		// Case #9: Ref3D. 3D reference to formula cell returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to formula cell returns TRUE. 1 of 1 argument used.');
		// Case #10: Area3D. 3D single-cell range with formula returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with formula returns TRUE. 1 of 1 argument used.');
		// Case #11: Table. Table column referencing formula cell returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column referencing formula cell returns TRUE. 1 of 1 argument used.');
		// Case #12: Formula. IF formula returning reference to formula cell returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(IF(TRUE,A103,A104))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(IF(TRUE,A103,A104)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. IF formula returning reference to formula cell returns TRUE. 1 of 1 argument used.');
		// Case #13: Formula. Formula returning reference to constant cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(INDIRECT("A103"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(INDIRECT("A103")) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Formula returning reference to constant cell returns FALSE. 1 of 1 argument used.');
		// Case #14: Reference link. Reference link to formula cell (A100=A1) returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference link to formula cell (A100=A1) returns TRUE. 1 of 1 argument used.');
		// Case #15: Name. Named range referencing constant cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range referencing constant cell returns FALSE. 1 of 1 argument used.');
		// Case #16: Ref3D. 3D reference to constant cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to constant cell returns FALSE. 1 of 1 argument used.');
		// Case #17: Formula. CHOOSE formula returning reference to formula cell returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(CHOOSE(1,A100))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(CHOOSE(1,A100)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. CHOOSE formula returning reference to formula cell returns TRUE. 1 of 1 argument used.');
		// Case #18: Reference link. Reference link to constant cell (A101=A2) returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference link to constant cell (A101=A2) returns FALSE. 1 of 1 argument used.');
		// Case #19: Name. Named range referencing text cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range referencing text cell returns FALSE. 1 of 1 argument used.');
		// Case #20: Formula. ADDRESS formula returning reference to formula cell (A1) returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(ADDRESS(1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(ADDRESS(1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. ADDRESS formula returning reference to formula cell (A1) returns TRUE. 1 of 1 argument used.');
		// Case #21: Table. Table column referencing constant cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Table. Table column referencing constant cell returns FALSE. 1 of 1 argument used.');

		// Negative cases:
		// Case #1: Number. Numeric reference returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(123)', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA(123) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Numeric reference returns #VALUE!. 1 of 1 argument used.');
		// Case #2: String. Text reference returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA("text")', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA("text") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA("text")', 'Test: Negative case: String. Text reference returns #VALUE!. 1 of 1 argument used.');
		// Case #3: Boolean. Boolean reference returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(FALSE)', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA(FALSE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA(FALSE)', 'Test: Negative case: Boolean. Boolean reference returns #VALUE!. 1 of 1 argument used.');
		// Case #4: Error. Error reference returns #N/A. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(NA())', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA(NA()) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA(NA())', 'Test: Negative case: Error. Error reference returns #N/A. 1 of 1 argument used.');
		// Case #5: Area. Multi-cell range returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 of 1 argument used.');
		// Case #6: Reference link. Reference to text cell (A103="text") returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to text cell (A103="text") returns FALSE. 1 of 1 argument used.');
		// Case #7: Name. Named range with non-reference (text) returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with non-reference (text) returns #VALUE!. 1 of 1 argument used.');
		// Case #8: Name3D. 3D named range with non-reference (text) returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with non-reference (text) returns #VALUE!. 1 of 1 argument used.');
		// Case #9: Ref3D. 3D reference to text cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Ref3D. 3D reference to text cell returns FALSE. 1 of 1 argument used.');
		// Case #10: Area3D. 3D multi-cell range returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 of 1 argument used.');
		// Case #11: Table. Table column with text value returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Table. Table column with text value returns FALSE. 1 of 1 argument used.');
		// Case #12: Formula. Formula returning #NUM! returns #NUM!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(SQRT(-1))', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA(SQRT(-1)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA(SQRT(-1))', 'Test: Negative case: Formula. Formula returning #NUM! returns #NUM!. 1 of 1 argument used.');
		// Case #13: Array. Array literal returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA({"A1"})', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA({"A1"}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA({"A1"})', 'Test: Negative case: Array. Array literal returns #VALUE!. 1 of 1 argument used.');
		// Case #14: Formula. Formula returning text ("A1") returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(CONCATENATE("A","1"))', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA(CONCATENATE("A","1")) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA(CONCATENATE("A";"1"))', 'Test: Negative case: Formula. Formula returning text ("A1") returns #VALUE!. 1 of 1 argument used.');
		// Case #15: Empty. Missing reference returns #VALUE!. 0 of 1 argument used.
		oParser = new parserFormula('ISFORMULA("")', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA("") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA()', 'Test: Negative case: Empty. Missing reference returns #VALUE!. 0 of 1 argument used.');
		// Case #16: Date. Date serial number returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(DATE(2025,1,1))', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA(DATE(2025,1,1)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA(DATE(2025;1;1))', 'Test: Negative case: Date. Date serial number returns #VALUE!. 1 of 1 argument used.');
		// Case #17: Time. Time value returns #VALUE!. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TIME(12,0,0))', 'A2', ws);
		//? assert.ok(oParser.parse() === false, 'Test: ISFORMULA(TIME(12,0,0)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '=ISFORMULA(TIME(12;0;0))', 'Test: Negative case: Time. Time value returns #VALUE!. 1 of 1 argument used.');
		// Case #18: Reference link. Reference to non-existent cell (A5) may return FALSE or error (context-dependent). 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to non-existent cell (A5) may return FALSE or error (context-dependent). 1 of 1 argument used.');
		// Case #19: Name. Named range referencing empty cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TestName3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(TestName3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range referencing empty cell returns FALSE. 1 of 1 argument used.');
		// Case #20: Formula. Formula returning reference to text cell returns FALSE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(INDIRECT("A103"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(INDIRECT("A103")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Formula returning reference to text cell returns FALSE. 1 of 1 argument used.');

		// Bounded cases:
		// Case #1: Reference link. Reference to first cell (A1, with formula) returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Reference link. Reference to first cell (A1, with formula) returns TRUE. 1 of 1 argument used.');
		// Case #2: Reference link. Reference to last cell in Excel (XFD1048576) returns FALSE (assuming no formula). 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(XFD1048576)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(XFD1048576) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Reference link. Reference to last cell in Excel (XFD1048576) returns FALSE (assuming no formula). 1 of 1 argument used.');
		// Case #3: Name. Named range referencing single formula cell (A1) returns TRUE. 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Name. Named range referencing single formula cell (A1) returns TRUE. 1 of 1 argument used.');
		// Case #4: Formula. Formula returning reference to last cell returns FALSE (assuming no formula). 1 of 1 argument used.
		oParser = new parserFormula('ISFORMULA(INDIRECT("XFD1048576"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISFORMULA(INDIRECT("XFD1048576")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Formula returning reference to last cell returns FALSE (assuming no formula). 1 of 1 argument used.');

		// TODO formula can't be enter without correct reference to a cell or name
		// Need to fix: parse error
		// Case #5: Reference link. Reference to empty cell returns FALSE. 1 of 1 argument used.
		// Case #12: Formula. IF formula returning reference to formula cell returns TRUE. 1 of 1 argument used.
		// Case #13: Formula. Formula returning reference to constant cell returns FALSE. 1 of 1 argument used.
		// Case #1: Number. Numeric reference returns #VALUE!. 1 of 1 argument used. - formula should not be parsed
		// Case #2: String. Text reference returns #VALUE!. 1 of 1 argument used.
		// Case #3: Boolean. Boolean reference returns #VALUE!. 1 of 1 argument used.
		// Case #4: Error. Error reference returns #N/A. 1 of 1 argument used.
		// Case #12: Formula. Formula returning #NUM! returns #NUM!. 1 of 1 argument used.
		// Case #13: Array. Array literal returns #VALUE!. 1 of 1 argument used.
		// Case #14: Formula. Formula returning text ("A1") returns #VALUE!. 1 of 1 argument used.
		// Case #15: Empty. Missing reference returns #VALUE!. 0 of 1 argument used.
		// Case #16: Date. Date serial number returns #VALUE!. 1 of 1 argument used.
		// Case #17: Time. Time value returns #VALUE!. 1 of 1 argument used.

	});

    QUnit.test("Test: \"ISLOGICAL\"", function (assert) {

		oParser = new parserFormula('ISLOGICAL(TRUE)', "A1", ws);
		assert.ok(oParser.parse(), 'ISLOGICAL(TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'ISLOGICAL(TRUE)');

		oParser = new parserFormula('ISLOGICAL("TRUE")', "A1", ws);
		assert.ok(oParser.parse(), 'ISLOGICAL("TRUE")');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'ISLOGICAL("TRUE")');

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("FALSE");
		// For area
		ws.getRange2("A102").setValue("123s");
		ws.getRange2("A103").setValue("TRUE");
		ws.getRange2("A105").setValue("");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		ws.getRange2("C601").setValue("TRUE"); // Bool (Column3)
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
		// Case #1: Boolean. Boolean TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Boolean TRUE, returns TRUE');
		// Case #2: Boolean. Boolean FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Boolean FALSE, returns TRUE');
		// Case #3: Formula. Formula evaluating to TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(1=1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(1=1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Formula evaluating to TRUE, returns TRUE');
		// Case #4: Formula. Formula evaluating to FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(1>2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(1>2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Formula evaluating to FALSE, returns TRUE');
		// Case #5: Reference link. Reference to cell with TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with TRUE, returns TRUE');
		// Case #6: Reference link. Reference to cell with FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with FALSE, returns TRUE');
		// Case #7: Area. Single-cell range with TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell range with TRUE, returns TRUE');
		// Case #8: Area. Multi-cell range with FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Multi-cell range with FALSE, returns TRUE');
		// Case #9: Array. Array with single boolean, returns TRUE
		oParser = new parserFormula('ISLOGICAL({TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL({TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with single boolean, returns TRUE');
		// Case #10: Array. Array with multiple booleans, evaluates first (FALSE), returns TRUE
		oParser = new parserFormula('ISLOGICAL({FALSE,TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL({FALSE,TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with multiple booleans, evaluates first (FALSE), returns TRUE');
		// Case #11: Name. Named range with TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with TRUE, returns TRUE');
		// Case #12: Name3D. 3D named range with FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with FALSE, returns TRUE');
		// Case #13: Ref3D. 3D reference to cell with TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with TRUE, returns TRUE');
		// Case #14: Area3D. 3D single-cell range with FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with FALSE, returns TRUE');
		// Case #15: Table. Table column with TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with TRUE, returns TRUE');
		// Case #16: Formula. Nested IF returning TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(IF(1=1,TRUE,FALSE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(IF(1=1,TRUE,FALSE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF returning TRUE, returns TRUE');
		// Case #17: Formula. Nested AND formula returning FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(AND(TRUE,FALSE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(AND(TRUE,FALSE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested AND formula returning FALSE, returns TRUE');
		// Case #18: Reference link. Reference to cell with formula evaluating to TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with formula evaluating to TRUE, returns TRUE');
		// Case #19: Name. Named range with FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with FALSE, returns TRUE');
		// Case #20: Table. Table column with FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with FALSE, returns TRUE');
		// Case #21: Formula. Nested OR formula returning TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(OR(1>2,2>1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(OR(1>2,2>1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested OR formula returning TRUE, returns TRUE');
		// Case #22: Ref3D. 3D reference to cell with formula evaluating to FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with formula evaluating to FALSE, returns TRUE');
		// Case #23: Area3D. 3D single-cell range with TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(Sheet2!A4:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Sheet2!A4:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with TRUE, returns TRUE');
		// Case #24: Name3D. 3D named range with TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with TRUE, returns TRUE');
		// Case #25: Formula. Nested NOT formula returning FALSE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(NOT(TRUE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(NOT(TRUE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested NOT formula returning FALSE, returns TRUE');

		// Negative cases:
		// Case #1: Number. Number input, returns FALSE
		oParser = new parserFormula('ISLOGICAL(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Number input, returns FALSE');
		// Case #2: Number. Zero, returns FALSE
		oParser = new parserFormula('ISLOGICAL(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Zero, returns FALSE');
		// Case #3: String. Text string, returns FALSE
		oParser = new parserFormula('ISLOGICAL("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Text string, returns FALSE');
		// Case #4: String. Empty string, returns FALSE
		oParser = new parserFormula('ISLOGICAL("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Empty string, returns FALSE');
		// Case #5: Empty. Empty argument, returns FALSE
		oParser = new parserFormula('ISLOGICAL(#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Empty. Empty argument, returns FALSE');
		// Case #6: Error. Error (#DIV/0!), returns FALSE
		oParser = new parserFormula('ISLOGICAL(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Error. Error (#DIV/0!), returns FALSE');
		// Case #7: Error. Error (#N/A), returns FALSE
		oParser = new parserFormula('ISLOGICAL(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Error. Error (#N/A), returns FALSE');
		// Case #8: Date. Date (serial number), returns FALSE
		oParser = new parserFormula('ISLOGICAL(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Date. Date (serial number), returns FALSE');
		// Case #9: Time. Time (serial number), returns FALSE
		oParser = new parserFormula('ISLOGICAL(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Time. Time (serial number), returns FALSE');
		// Case #10: Reference link. Reference to cell with number, returns FALSE
		oParser = new parserFormula('ISLOGICAL(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with number, returns FALSE');
		// Case #11: Reference link. Reference to cell with text, returns FALSE
		oParser = new parserFormula('ISLOGICAL(A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with text, returns FALSE');
		// Case #12: Area. Single-cell range with number, returns FALSE
		oParser = new parserFormula('ISLOGICAL(A108:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A108:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Single-cell range with number, returns FALSE');
		// Case #13: Area. Multi-cell range with text, returns FALSE
		oParser = new parserFormula('ISLOGICAL(A109:A110)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A109:A110) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range with text, returns FALSE');
		// Case #14: Array. Array with number, returns FALSE
		oParser = new parserFormula('ISLOGICAL({123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL({123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with number, returns FALSE');
		// Case #15: Name. Named range with error, returns FALSE
		oParser = new parserFormula('ISLOGICAL(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with error, returns FALSE');
		// Case #16: Name3D. 3D named range with text, returns FALSE
		oParser = new parserFormula('ISLOGICAL(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with text, returns FALSE');
		// Case #17: Ref3D. 3D reference to cell with number, returns FALSE
		oParser = new parserFormula('ISLOGICAL(Sheet2!A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Sheet2!A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Ref3D. 3D reference to cell with number, returns FALSE');
		// Case #18: Area3D. 3D multi-cell range with error, returns FALSE
		oParser = new parserFormula('ISLOGICAL(Sheet2!A6:A7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Sheet2!A6:A7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range with error, returns FALSE');
		// Case #19: Table. Table column with text, returns FALSE
		oParser = new parserFormula('ISLOGICAL(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Table. Table column with text, returns FALSE');
		// Case #20: Formula. Nested SQRT formula returning number, returns FALSE
		oParser = new parserFormula('ISLOGICAL(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested SQRT formula returning number, returns FALSE');

		// Bounded cases:
		// Case #1: Formula. Simple logical expression (TRUE), returns TRUE
		oParser = new parserFormula('ISLOGICAL(1=1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(1=1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Simple logical expression (TRUE), returns TRUE');
		// Case #2: Formula. Simple logical expression (FALSE), returns TRUE
		oParser = new parserFormula('ISLOGICAL(1<>1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(1<>1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Simple logical expression (FALSE), returns TRUE');
		// Case #3: Reference link. Reference to cell with complex logical formula, returns TRUE
		oParser = new parserFormula('ISLOGICAL(A111)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(A111) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Reference link. Reference to cell with complex logical formula, returns TRUE');
		// Case #4: Formula. Nested IFERROR returning TRUE, returns TRUE
		oParser = new parserFormula('ISLOGICAL(IFERROR(1/0,TRUE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISLOGICAL(IFERROR(1/0,TRUE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Nested IFERROR returning TRUE, returns TRUE');

		// Need to fix:
		// Case #8: Area. Multi-cell range with FALSE, returns TRUE

		testArrayFormula2(assert, "ISLOGICAL", 1, 1);
	});

    QUnit.test("Test: \"ISNA\"", function (assert) {
		ws.getRange2("A1").setValue("#N/A");

		oParser = new parserFormula("ISNA(A1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("#DIV/0!");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Error. #DIV/0! error, returns TRUE
		oParser = new parserFormula('ISNA(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. #DIV/0! error, returns TRUE');
		// Case #2: Error. #VALUE! error, returns TRUE
		oParser = new parserFormula('ISNA(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. #VALUE! error, returns TRUE');
		// Case #3: Error. #REF! error, returns TRUE
		oParser = new parserFormula('ISNA(#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. #REF! error, returns TRUE');
		// Case #4: Error. #NAME? error, returns TRUE
		oParser = new parserFormula('ISNA(#NAME?)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(#NAME?) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. #NAME? error, returns TRUE');
		// Case #5: Error. #NUM! error, returns TRUE
		oParser = new parserFormula('ISNA(#NUM!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(#NUM!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. #NUM! error, returns TRUE');
		// Case #6: Error. #NULL! error, returns TRUE
		oParser = new parserFormula('ISNA(#NULL!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(#NULL!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. #NULL! error, returns TRUE');
		// Case #7: Formula. Formula resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNA(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Formula resulting in #DIV/0!, returns TRUE');
		// Case #8: Formula. Nested formula resulting in #NUM!, returns TRUE
		oParser = new parserFormula('ISNA(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested formula resulting in #NUM!, returns TRUE');
		// Case #9: Reference link. Reference to cell with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNA(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with #DIV/0!, returns TRUE');
		// Case #10: Reference link. Reference to cell with #VALUE!, returns TRUE
		oParser = new parserFormula('ISNA(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with #VALUE!, returns TRUE');
		// Case #11: Area. Single-cell range with #REF!, returns TRUE
		oParser = new parserFormula('ISNA(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell range with #REF!, returns TRUE');
		// Case #12: Area. Multi-cell range with error, returns TRUE
		oParser = new parserFormula('ISNA(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Multi-cell range with error, returns TRUE');
		// Case #13: Array. Array with error value, returns TRUE
		oParser = new parserFormula('ISNA({#DIV/0!})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA({#DIV/0!}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with error value, returns TRUE');
		// Case #14: Name. Named range with #NUM!, returns TRUE
		oParser = new parserFormula('ISNA(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with #NUM!, returns TRUE');
		// Case #15: Name3D. 3D named range with #NULL!, returns TRUE
		oParser = new parserFormula('ISNA(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with #NULL!, returns TRUE');
		// Case #16: Ref3D. 3D reference to cell with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNA(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with #DIV/0!, returns TRUE');
		// Case #17: Area3D. 3D single-cell range with #VALUE!, returns TRUE
		oParser = new parserFormula('ISNA(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with #VALUE!, returns TRUE');
		// Case #18: Table. Table column with #REF!, returns TRUE
		oParser = new parserFormula('ISNA(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with #REF!, returns TRUE');
		// Case #19: Formula. Nested IF resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNA(IF(TRUE,1/0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(IF(TRUE,1/0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested IF resulting in #DIV/0!, returns TRUE');
		// Case #20: Formula. Formula causing #VALUE!, returns TRUE
		oParser = new parserFormula('ISNA("text"+1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA("text"+1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Formula causing #VALUE!, returns TRUE');
		// Case #21: Reference link. Reference to cell with formula causing #NAME?, returns TRUE
		oParser = new parserFormula('ISNA(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with formula causing #NAME?, returns TRUE');
		// Case #22: Area3D. 3D multi-cell range with #NUM!, returns TRUE
		oParser = new parserFormula('ISNA(Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D multi-cell range with #NUM!, returns TRUE');
		// Case #23: Formula. Nested LN formula causing #NUM!, returns TRUE
		oParser = new parserFormula('ISNA(LN(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(LN(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested LN formula causing #NUM!, returns TRUE');
		// Case #24: Name. Named range with #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNA(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with #DIV/0!, returns TRUE');
		// Case #25: Table. Table column with formula causing #VALUE!, returns TRUE
		oParser = new parserFormula('ISNA(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with formula causing #VALUE!, returns TRUE');

		// Negative cases:
		// Case #1: Error. #N/A error, returns FALSE
		oParser = new parserFormula('ISNA(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Error. #N/A error, returns FALSE');
		// Case #2: Number. Number input, returns FALSE
		oParser = new parserFormula('ISNA(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Number input, returns FALSE');
		// Case #3: String. String input, returns FALSE
		oParser = new parserFormula('ISNA("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. String input, returns FALSE');
		// Case #4: Empty. Empty argument, returns FALSE
		oParser = new parserFormula('ISNA("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Empty. Empty argument, returns FALSE');
		// Case #5: Boolean. Boolean input, returns FALSE
		oParser = new parserFormula('ISNA(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean input, returns FALSE');
		// Case #6: Boolean. Boolean input, returns FALSE
		oParser = new parserFormula('ISNA(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean input, returns FALSE');
		// Case #7: Formula. Date formula returns serial number, returns FALSE
		oParser = new parserFormula('ISNA(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Date formula returns serial number, returns FALSE');
		// Case #8: Formula. Time formula returns serial number, returns FALSE
		oParser = new parserFormula('ISNA(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Time formula returns serial number, returns FALSE');
		// Case #9: Reference link. Reference to empty cell, returns FALSE
		oParser = new parserFormula('ISNA(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to empty cell, returns FALSE');
		// Case #10: Reference link. Reference to cell with number, returns FALSE
		oParser = new parserFormula('ISNA(A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with number, returns FALSE');
		// Case #11: Area. Single-cell range with string, returns FALSE
		oParser = new parserFormula('ISNA(A108:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A108:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Single-cell range with string, returns FALSE');
		// Case #12: Area. Multi-cell range with non-error values, returns FALSE
		oParser = new parserFormula('ISNA(A109:A110)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(A109:A110) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range with non-error values, returns FALSE');
		// Case #13: Array. Array with number, returns FALSE
		oParser = new parserFormula('ISNA({123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA({123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with number, returns FALSE');
		// Case #14: Name. Named range with empty cell, returns FALSE
		oParser = new parserFormula('ISNA(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with empty cell, returns FALSE');
		// Case #15: Name3D. 3D named range with number, returns FALSE
		oParser = new parserFormula('ISNA(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with number, returns FALSE');
		// Case #16: Ref3D. 3D reference to cell with string, returns FALSE
		oParser = new parserFormula('ISNA(Sheet2!A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(Sheet2!A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Ref3D. 3D reference to cell with string, returns FALSE');
		// Case #17: Area3D. 3D multi-cell range with non-error values, returns FALSE
		oParser = new parserFormula('ISNA(Sheet2!A6:A7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(Sheet2!A6:A7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range with non-error values, returns FALSE');
		// Case #18: Table. Table column with boolean, returns FALSE
		oParser = new parserFormula('ISNA(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Table. Table column with boolean, returns FALSE');
		// Case #19: Formula. Nested IF returning number, returns FALSE
		oParser = new parserFormula('ISNA(IF(TRUE,123,NA()))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(IF(TRUE,123,NA())) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested IF returning number, returns FALSE');
		// Case #20: Formula. Nested SQRT formula returning number, returns FALSE
		oParser = new parserFormula('ISNA(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested SQRT formula returning number, returns FALSE');

		// Bounded cases:
		// Case #1: Formula. Minimum Excel date divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNA(DATE(1900,1,1)/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(DATE(1900,1,1)/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Minimum Excel date divided by zero, returns #DIV/0!, returns TRUE');
		// Case #2: Formula. Maximum Excel date divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNA(DATE(9999,12,31)/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(DATE(9999,12,31)/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Maximum Excel date divided by zero, returns #DIV/0!, returns TRUE');
		// Case #3: Formula. Smallest positive Excel number divided by zero, returns #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNA(1E-307/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(1E-307/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Smallest positive Excel number divided by zero, returns #DIV/0!, returns TRUE');
		// Case #4: Formula. Largest Excel number causing overflow (#NUM!), returns TRUE
		oParser = new parserFormula('ISNA(9.99999999999999E+307*10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNA(9.99999999999999E+307*10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Largest Excel number causing overflow (#NUM!), returns TRUE');


		testArrayFormula2(assert, "ISNA", 1, 1);
	});

    QUnit.test("Test: \"ISNONTEXT\"", function (assert) {
		oParser = new parserFormula('ISNONTEXT("123")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Number. Number input, returns TRUE
		oParser = new parserFormula('ISNONTEXT(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Number input, returns TRUE');
		// Case #2: Number. Zero, returns TRUE
		oParser = new parserFormula('ISNONTEXT(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Zero, returns TRUE');
		// Case #3: Empty. Empty argument (blank cell), returns TRUE
		oParser = new parserFormula('ISNONTEXT("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Empty. Empty argument (blank cell), returns TRUE');
		// Case #4: Error. Error (#DIV/0!), returns TRUE
		oParser = new parserFormula('ISNONTEXT(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. Error (#DIV/0!), returns TRUE');
		// Case #5: Error. Error (#VALUE!), returns TRUE
		oParser = new parserFormula('ISNONTEXT(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. Error (#VALUE!), returns TRUE');
		// Case #6: Error. Error (#N/A), returns TRUE
		oParser = new parserFormula('ISNONTEXT(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Error. Error (#N/A), returns TRUE');
		// Case #7: Boolean. Boolean TRUE, returns TRUE
		oParser = new parserFormula('ISNONTEXT(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Boolean TRUE, returns TRUE');
		// Case #8: Boolean. Boolean FALSE, returns TRUE
		oParser = new parserFormula('ISNONTEXT(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Boolean FALSE, returns TRUE');
		// Case #9: Formula. Date formula (serial number), returns TRUE
		oParser = new parserFormula('ISNONTEXT(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Date formula (serial number), returns TRUE');
		// Case #10: Formula. Time formula (serial number), returns TRUE
		oParser = new parserFormula('ISNONTEXT(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Time formula (serial number), returns TRUE');
		// Case #11: Reference link. Reference to empty cell, returns TRUE
		oParser = new parserFormula('ISNONTEXT(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to empty cell, returns TRUE');
		// Case #12: Reference link. Reference to cell with number, returns TRUE
		oParser = new parserFormula('ISNONTEXT(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with number, returns TRUE');
		// Case #13: Area. Single-cell range with number, returns TRUE
		oParser = new parserFormula('ISNONTEXT(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell range with number, returns TRUE');
		// Case #14: Area. Multi-cell range with number, returns TRUE
		oParser = new parserFormula('ISNONTEXT(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Multi-cell range with number, returns TRUE');
		// Case #15: Array. Array with number, returns TRUE
		oParser = new parserFormula('ISNONTEXT({123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT({123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with number, returns TRUE');
		// Case #16: Name. Named range with error, returns TRUE
		oParser = new parserFormula('ISNONTEXT(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name. Named range with error, returns TRUE');
		// Case #17: Name3D. 3D named range with boolean, returns TRUE
		oParser = new parserFormula('ISNONTEXT(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D. 3D named range with boolean, returns TRUE');
		// Case #18: Ref3D. 3D reference to cell with number, returns TRUE
		oParser = new parserFormula('ISNONTEXT(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Ref3D. 3D reference to cell with number, returns TRUE');
		// Case #19: Area3D. 3D single-cell range with error, returns TRUE
		oParser = new parserFormula('ISNONTEXT(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area3D. 3D single-cell range with error, returns TRUE');
		// Case #20: Table. Table column with empty cell, returns TRUE
		oParser = new parserFormula('ISNONTEXT(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Table. Table column with empty cell, returns TRUE');
		// Case #21: Formula. Nested IF returning number, returns TRUE
		oParser = new parserFormula('ISNONTEXT(IF(TRUE,123,"text"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(IF(TRUE,123,"text")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF returning number, returns TRUE');
		// Case #22: Formula. Nested SQRT formula returning number, returns TRUE
		oParser = new parserFormula('ISNONTEXT(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested SQRT formula returning number, returns TRUE');
		// Case #23: Time. Time (0.0, number), returns TRUE
		oParser = new parserFormula('ISNONTEXT(TIME(0,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TIME(0,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Time. Time (0.0, number), returns TRUE');
		// Case #24: Formula. Volatile NOW formula (serial number), returns TRUE
		oParser = new parserFormula('ISNONTEXT(NOW())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(NOW()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Volatile NOW formula (serial number), returns TRUE');
		// Case #25: Reference link. Reference to cell with boolean, returns TRUE
		oParser = new parserFormula('ISNONTEXT(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with boolean, returns TRUE');

		// Negative cases:
		// Case #1: String. Text string, returns FALSE
		oParser = new parserFormula('ISNONTEXT("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Text string, returns FALSE');
		// Case #2: String. Empty string, returns FALSE
		oParser = new parserFormula('ISNONTEXT("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Empty string, returns FALSE');
		// Case #3: String. Numeric string (still text), returns FALSE
		oParser = new parserFormula('ISNONTEXT("123")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT("123") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Numeric string (still text), returns FALSE');
		// Case #4: String. String with space, returns FALSE
		oParser = new parserFormula('ISNONTEXT(" ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(" ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. String with space, returns FALSE');
		// Case #5: Formula. Nested TEXT formula returning string, returns FALSE
		oParser = new parserFormula('ISNONTEXT(TEXT(123,"0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TEXT(123,"0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested TEXT formula returning string, returns FALSE');
		// Case #6: Reference link. Reference to cell with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with text, returns FALSE');
		// Case #7: Reference link. Reference to cell with empty string, returns FALSE
		oParser = new parserFormula('ISNONTEXT(A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with empty string, returns FALSE');
		// Case #8: Area. Single-cell range with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(A108:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A108:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Single-cell range with text, returns FALSE');
		// Case #9: Area. Multi-cell range with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(A109:A110)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A109:A110) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Multi-cell range with text, returns FALSE');
		// Case #10: Array. Array with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT({"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT({"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with text, returns FALSE');
		// Case #11: Array. Array with text as first element, returns FALSE
		oParser = new parserFormula('ISNONTEXT({"abc",123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT({"abc",123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with text as first element, returns FALSE');
		// Case #12: Name. Named range with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name. Named range with text, returns FALSE');
		// Case #13: Name. Named range with empty string, returns FALSE
		oParser = new parserFormula('ISNONTEXT(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name. Named range with empty string, returns FALSE');
		// Case #14: Name3D. 3D named range with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name3D. 3D named range with text, returns FALSE');
		// Case #15: Ref3D. 3D reference to cell with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Ref3D. 3D reference to cell with text, returns FALSE');
		// Case #16: Area3D. 3D multi-cell range with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area3D. 3D multi-cell range with text, returns FALSE');
		// Case #17: Table. Table column with text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Table. Table column with text, returns FALSE');
		// Case #18: Formula. Nested IF returning text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(IF(TRUE,"text",123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(IF(TRUE,"text",123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested IF returning text, returns FALSE');
		// Case #19: Formula. Nested CONCAT formula returning text, returns FALSE
		oParser = new parserFormula('ISNONTEXT(CONCAT("a","b"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(CONCAT("a","b")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested CONCAT formula returning text, returns FALSE');
		// Case #20: Reference link. Reference to cell with numeric string, returns FALSE
		oParser = new parserFormula('ISNONTEXT(A111)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(A111) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with numeric string, returns FALSE');

		// Bounded cases:
		// Case #1: Number. Smallest positive Excel number, returns TRUE
		oParser = new parserFormula('ISNONTEXT(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Smallest positive Excel number, returns TRUE');
		// Case #2: Number. Largest Excel number, returns TRUE
		oParser = new parserFormula('ISNONTEXT(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Largest Excel number, returns TRUE');
		// Case #3: Date. Minimum Excel date (serial number 1), returns TRUE
		oParser = new parserFormula('ISNONTEXT(DATE(1900,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(DATE(1900,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Date. Minimum Excel date (serial number 1), returns TRUE');
		// Case #4: Date. Maximum Excel date (serial number 2958465), returns TRUE
		oParser = new parserFormula('ISNONTEXT(DATE(9999,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(DATE(9999,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Date. Maximum Excel date (serial number 2958465), returns TRUE');
		// Case #5: Formula. Formula resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISNONTEXT(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNONTEXT(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Formula resulting in #DIV/0!, returns TRUE');

		// Need to fix:
		// Case #14: Area. Multi-cell range with number, returns TRUE

		testArrayFormula2(assert, "ISNONTEXT", 1, 1);
	});

	QUnit.test("Test: \"ISNUMBER\"", function (assert) {
		ws.getRange2("A1").setValue("123");

		oParser = new parserFormula('ISNUMBER(4)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('ISNUMBER(A1)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Number. Integer input, returns TRUE
		oParser = new parserFormula('ISNUMBER(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Integer input, returns TRUE');
		// Case #2: Number. Zero, returns TRUE
		oParser = new parserFormula('ISNUMBER(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Zero, returns TRUE');
		// Case #3: Number. Floating-point number, returns TRUE
		oParser = new parserFormula('ISNUMBER(3.14)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(3.14) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Floating-point number, returns TRUE');
		// Case #4: Date. Date (serial number), returns TRUE
		oParser = new parserFormula('ISNUMBER(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Date. Date (serial number), returns TRUE');
		// Case #5: Time. Time (fractional number), returns TRUE
		oParser = new parserFormula('ISNUMBER(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Time. Time (fractional number), returns TRUE');
		// Case #6: Formula. Nested SQRT formula returning number, returns TRUE
		oParser = new parserFormula('ISNUMBER(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested SQRT formula returning number, returns TRUE');
		// Case #7: Formula. Nested ROUND formula returning number, returns TRUE
		oParser = new parserFormula('ISNUMBER(ROUND(2.718,2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(ROUND(2.718,2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested ROUND formula returning number, returns TRUE');
		// Case #8: Reference link. Reference to cell with number, returns TRUE
		oParser = new parserFormula('ISNUMBER(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with number, returns TRUE');
		// Case #9: Reference link. Reference to cell with float, returns TRUE
		oParser = new parserFormula('ISNUMBER(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with float, returns TRUE');
		// Case #10: Area. Single-cell range with number, returns TRUE
		oParser = new parserFormula('ISNUMBER(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell range with number, returns TRUE');
		// Case #11: Area. Multi-cell range with number, returns TRUE
		oParser = new parserFormula('ISNUMBER(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Multi-cell range with number, returns TRUE');
		// Case #12: Array. Array with single number, returns TRUE
		oParser = new parserFormula('ISNUMBER({123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER({123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with single number, returns TRUE');
		// Case #13: Array. Array with multiple numbers, evaluates first, returns TRUE
		oParser = new parserFormula('ISNUMBER({456,789})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER({456,789}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with multiple numbers, evaluates first, returns TRUE');
		// Case #14: Name. Named range with number, returns TRUE
		oParser = new parserFormula('ISNUMBER(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name. Named range with number, returns TRUE');
		// Case #15: Name3D. 3D named range with number, returns TRUE
		oParser = new parserFormula('ISNUMBER(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D. 3D named range with number, returns TRUE');
		// Case #16: Ref3D. 3D reference to cell with number, returns TRUE
		oParser = new parserFormula('ISNUMBER(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Ref3D. 3D reference to cell with number, returns TRUE');
		// Case #17: Area3D. 3D single-cell range with number, returns TRUE
		oParser = new parserFormula('ISNUMBER(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area3D. 3D single-cell range with number, returns TRUE');
		// Case #18: Table. Table column with number, returns TRUE
		oParser = new parserFormula('ISNUMBER(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Table. Table column with number, returns TRUE');
		// Case #19: Formula. Nested IF returning number, returns TRUE
		oParser = new parserFormula('ISNUMBER(IF(TRUE,123,456))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(IF(TRUE,123,456)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF returning number, returns TRUE');
		// Case #20: Formula. Volatile NOW formula (serial number), returns TRUE
		oParser = new parserFormula('ISNUMBER(NOW())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(NOW()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Volatile NOW formula (serial number), returns TRUE');
		// Case #21: String. Numeric string converted to number, returns TRUE
		oParser = new parserFormula('ISNUMBER("123")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER("123") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String. Numeric string converted to number, returns TRUE');
		// Case #22: Formula. Nested VALUE formula converting string to number, returns TRUE
		oParser = new parserFormula('ISNUMBER(VALUE("456"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(VALUE("456")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested VALUE formula converting string to number, returns TRUE');
		// Case #23: Reference link. Reference to cell with date serial number, returns TRUE
		oParser = new parserFormula('ISNUMBER(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with date serial number, returns TRUE');
		// Case #24: Table. Table column with float, returns TRUE
		oParser = new parserFormula('ISNUMBER(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with float, returns TRUE');
		// Case #25: Name3D. 3D named range with float, returns TRUE
		oParser = new parserFormula('ISNUMBER(TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with float, returns TRUE');

		// Negative cases:
		// Case #1: String. Non-numeric text string, returns FALSE
		oParser = new parserFormula('ISNUMBER("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Non-numeric text string, returns FALSE');
		// Case #2: String. Empty string, returns FALSE
		oParser = new parserFormula('ISNUMBER("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Empty string, returns FALSE');
		// Case #3: Boolean. Boolean TRUE, returns FALSE
		oParser = new parserFormula('ISNUMBER(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean TRUE, returns FALSE');
		// Case #4: Boolean. Boolean FALSE, returns FALSE
		oParser = new parserFormula('ISNUMBER(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean FALSE, returns FALSE');
		// Case #5: Empty. Empty argument, returns FALSE
		oParser = new parserFormula('ISNUMBER(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Empty. Empty argument, returns FALSE');
		// Case #6: Error. Error (#DIV/0!), returns FALSE
		oParser = new parserFormula('ISNUMBER(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Error. Error (#DIV/0!), returns FALSE');
		// Case #7: Error. Error (#N/A), returns FALSE
		oParser = new parserFormula('ISNUMBER(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Error. Error (#N/A), returns FALSE');
		// Case #8: Formula. Formula causing #VALUE!, returns FALSE
		oParser = new parserFormula('ISNUMBER("text"+1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER("text"+1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Formula causing #VALUE!, returns FALSE');
		// Case #9: Reference link. Reference to cell with text, returns FALSE
		oParser = new parserFormula('ISNUMBER(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with text, returns FALSE');
		// Case #10: Reference link. Reference to cell with boolean, returns FALSE
		oParser = new parserFormula('ISNUMBER(A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with boolean, returns FALSE');
		// Case #11: Area. Single-cell range with text, returns FALSE
		oParser = new parserFormula('ISNUMBER(A108:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A108:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Single-cell range with text, returns FALSE');
		// Case #12: Area. Multi-cell range with error, returns FALSE
		oParser = new parserFormula('ISNUMBER(A109:A110)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(A109:A110) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range with error, returns FALSE');
		// Case #13: Array. Array with text, returns FALSE
		oParser = new parserFormula('ISNUMBER({"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER({"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with text, returns FALSE');
		// Case #14: Name. Named range with text, returns FALSE
		oParser = new parserFormula('ISNUMBER(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name. Named range with text, returns FALSE');
		// Case #15: Name3D. 3D named range with text, returns FALSE
		oParser = new parserFormula('ISNUMBER(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name3D. 3D named range with text, returns FALSE');
		// Case #16: Ref3D. 3D reference to cell with error, returns FALSE
		oParser = new parserFormula('ISNUMBER(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Ref3D. 3D reference to cell with error, returns FALSE');
		// Case #17: Area3D. 3D multi-cell range with text, returns FALSE
		oParser = new parserFormula('ISNUMBER(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range with text, returns FALSE');
		// Case #18: Table. Table column with boolean, returns FALSE
		oParser = new parserFormula('ISNUMBER(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Table. Table column with boolean, returns FALSE');
		// Case #19: Formula. Nested IF returning text, returns FALSE
		oParser = new parserFormula('ISNUMBER(IF(TRUE,"text",123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(IF(TRUE,"text",123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested IF returning text, returns FALSE');
		// Case #20: Formula. Nested CONCAT formula returning text, returns FALSE
		oParser = new parserFormula('ISNUMBER(CONCAT("a","b"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(CONCAT("a","b")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested CONCAT formula returning text, returns FALSE');

		// Bounded cases:
		// Case #1: Number. Smallest positive Excel number, returns TRUE
		oParser = new parserFormula('ISNUMBER(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Smallest positive Excel number, returns TRUE');
		// Case #2: Number. Largest Excel number, returns TRUE
		oParser = new parserFormula('ISNUMBER(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Largest Excel number, returns TRUE');
		// Case #3: Date. Minimum Excel date (serial number 1), returns TRUE
		oParser = new parserFormula('ISNUMBER(DATE(1900,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(DATE(1900,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Date. Minimum Excel date (serial number 1), returns TRUE');
		// Case #4: Date. Maximum Excel date (serial number 2958465), returns TRUE
		oParser = new parserFormula('ISNUMBER(DATE(9999,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(DATE(9999,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Date. Maximum Excel date (serial number 2958465), returns TRUE');
		// Case #5: Formula. Largest negative Excel number, returns TRUE
		oParser = new parserFormula('ISNUMBER(-9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISNUMBER(-9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula. Largest negative Excel number, returns TRUE');


		testArrayFormula2(assert, "ISNUMBER", 1, 1);
	});

	QUnit.test("Test: \"ISODD\"", function (assert) {
		oParser = new parserFormula('ISODD(-1)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('ISODD(2.5)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula('ISODD(5)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Number. Integer input (even), returns TRUE
		oParser = new parserFormula('ISODD(2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Integer input (even), returns TRUE');
		// Case #2: Number. Integer input (odd), returns FALSE
		oParser = new parserFormula('ISODD(3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Integer input (odd), returns FALSE');
		// Case #3: Number. Zero (even), returns TRUE
		oParser = new parserFormula('ISODD(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Zero (even), returns TRUE');
		// Case #4: Number. Float truncated to 2 (even), returns TRUE
		oParser = new parserFormula('ISODD(2.7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(2.7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Float truncated to 2 (even), returns TRUE');
		// Case #5: Number. Float truncated to 3 (odd), returns FALSE
		oParser = new parserFormula('ISODD(3.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(3.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Float truncated to 3 (odd), returns FALSE');
		// Case #6: Formula. Nested ROUND formula (2, even), returns TRUE
		oParser = new parserFormula('ISODD(ROUND(2.3,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(ROUND(2.3,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested ROUND formula (2, even), returns TRUE');
		// Case #7: Formula. Nested SQRT formula (3, odd), returns FALSE
		oParser = new parserFormula('ISODD(SQRT(9))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(SQRT(9)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested SQRT formula (3, odd), returns FALSE');
		// Case #8: Reference link. Reference to cell with even number, returns TRUE
		oParser = new parserFormula('ISODD(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with even number, returns TRUE');
		// Case #9: Reference link. Reference to cell with odd number, returns FALSE
		oParser = new parserFormula('ISODD(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with odd number, returns FALSE');
		// Case #10: Area. Single-cell range with even number, returns TRUE
		oParser = new parserFormula('ISODD(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Single-cell range with even number, returns TRUE');
		// Case #11: Area. Multi-cell range with number, returns TRUE (first cell truncated to 2)
		oParser = new parserFormula('ISODD(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Multi-cell range with number, returns TRUE (first cell truncated to 2)');
		// Case #12: Array. Array with single even number, returns TRUE
		oParser = new parserFormula('ISODD({4})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD({4}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with single even number, returns TRUE');
		// Case #13: Array. Array with multiple numbers, evaluates first (3, odd), returns FALSE
		oParser = new parserFormula('ISODD({3,4})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD({3,4}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with multiple numbers, evaluates first (3, odd), returns FALSE');
		// Case #14: Name. Named range with even number, returns TRUE
		oParser = new parserFormula('ISODD(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with even number, returns TRUE');
		// Case #15: Name3D. 3D named range with odd number, returns FALSE
		oParser = new parserFormula('ISODD(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with odd number, returns FALSE');
		// Case #16: Ref3D. 3D reference to cell with even number, returns TRUE
		oParser = new parserFormula('ISODD(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Ref3D. 3D reference to cell with even number, returns TRUE');
		// Case #17: Area3D. 3D single-cell range with odd number, returns FALSE
		oParser = new parserFormula('ISODD(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with odd number, returns FALSE');
		// Case #18: Table. Table column with even number, returns TRUE
		oParser = new parserFormula('ISODD(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Table. Table column with even number, returns TRUE');
		// Case #19: Formula. Nested IF returning even number, returns TRUE
		oParser = new parserFormula('ISODD(IF(TRUE,4,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(IF(TRUE,4,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested IF returning even number, returns TRUE');
		// Case #20: Date. Date (serial number 45654, even), returns TRUE
		oParser = new parserFormula('ISODD(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Date. Date (serial number 45654, even), returns TRUE');
		// Case #21: Time. Time scaled to integer (500000, even), returns TRUE
		oParser = new parserFormula('ISODD(TIME(12,0,0)*1000000)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(TIME(12,0,0)*1000000) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Time. Time scaled to integer (500000, even), returns TRUE');
		// Case #22: String. String convertible to even number, returns TRUE
		oParser = new parserFormula('ISODD("4")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD("4") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String. String convertible to even number, returns TRUE');
		// Case #23: Formula. Volatile NOW formula (serial number, e.g., 45654), returns TRUE
		oParser = new parserFormula('ISODD(NOW())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(NOW()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Volatile NOW formula (serial number, e.g., 45654), returns TRUE');
		// Case #24: Name. Named range with float truncated to even, returns TRUE
		oParser = new parserFormula('ISODD(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with float truncated to even, returns TRUE');
		// Case #25: Table. Table column with odd number, returns FALSE
		oParser = new parserFormula('ISODD(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Table. Table column with odd number, returns FALSE');

		// Negative cases:
		// Case #1: Error. Error input (#DIV/0!), returns #VALUE!
		oParser = new parserFormula('ISODD(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Error. Error input (#DIV/0!), returns #VALUE!');
		// Case #2: Error. Error input (#VALUE!), returns #VALUE!
		oParser = new parserFormula('ISODD(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error. Error input (#VALUE!), returns #VALUE!');
		// Case #3: Error. Error input (#N/A), returns #VALUE!
		oParser = new parserFormula('ISODD(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error input (#N/A), returns #VALUE!');
		// Case #4: Error. Error input (#REF!), returns #VALUE!
		oParser = new parserFormula('ISODD(#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#REF!', 'Test: Negative case: Error. Error input (#REF!), returns #VALUE!');
		// Case #5: String. Non-numeric string, returns #VALUE!
		oParser = new parserFormula('ISODD("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string, returns #VALUE!');
		// Case #6: Empty. Empty argument, returns #VALUE!
		oParser = new parserFormula('ISODD("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty argument, returns #VALUE!');
		// Case #7: Boolean. Boolean input, returns #VALUE!
		oParser = new parserFormula('ISODD(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean input, returns #VALUE!');
		// Case #8: Boolean. Boolean input, returns #VALUE!
		oParser = new parserFormula('ISODD(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(FALSE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean input, returns #VALUE!');
		// Case #9: Formula. Formula causing #VALUE!, returns #VALUE!
		oParser = new parserFormula('ISODD("text"+1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD("text"+1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Formula causing #VALUE!, returns #VALUE!');
		// Case #10: Reference link. Reference to cell with error, returns #VALUE!
		oParser = new parserFormula('ISODD(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with error, returns #VALUE!');
		// Case #11: Reference link. Reference to cell with non-numeric string, returns #VALUE!
		oParser = new parserFormula('ISODD(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with non-numeric string, returns #VALUE!');
		// Case #12: Area. Single-cell range with error, returns #VALUE!
		oParser = new parserFormula('ISODD(A107:A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(A107:A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Single-cell range with error, returns #VALUE!');
		// Case #13: Area. Multi-cell range with non-numeric, returns #VALUE!
		oParser = new parserFormula('ISODD(A108:A109)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(A108:A109) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range with non-numeric, returns #VALUE!');
		// Case #14: Array. Array with non-numeric string, returns #VALUE!
		oParser = new parserFormula('ISODD({"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD({"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array. Array with non-numeric string, returns #VALUE!');
		// Case #15: Name. Named range with error, returns #VALUE!
		oParser = new parserFormula('ISODD(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with error, returns #VALUE!');
		// Case #16: Name3D. 3D named range with non-numeric string, returns #VALUE!
		oParser = new parserFormula('ISODD(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with non-numeric string, returns #VALUE!');
		// Case #17: Ref3D. 3D reference to cell with error, returns #VALUE!
		oParser = new parserFormula('ISODD(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to cell with error, returns #VALUE!');
		// Case #18: Area3D. 3D multi-cell range with non-numeric, returns #VALUE!
		oParser = new parserFormula('ISODD(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D multi-cell range with non-numeric, returns #VALUE!');
		// Case #19: Table. Table column with error, returns #VALUE!
		oParser = new parserFormula('ISODD(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with error, returns #VALUE!');
		// Case #20: Formula. Nested formula causing #NUM!, returns #VALUE!
		oParser = new parserFormula('ISODD(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Nested formula causing #NUM!, returns #VALUE!');

		// Bounded cases:
		// Case #1: Number. Smallest positive integer (odd), returns FALSE
		oParser = new parserFormula('ISODD(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Smallest positive integer (odd), returns FALSE');
		// Case #2: Number. Smallest negative integer (odd), returns FALSE
		oParser = new parserFormula('ISODD(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Smallest negative integer (odd), returns FALSE');
		// Case #3: Number. Largest even 32-bit integer, returns TRUE
		oParser = new parserFormula('ISODD(2^31-2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(2^31-2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Largest even 32-bit integer, returns TRUE');
		// Case #4: Number. Largest negative odd 32-bit integer, returns FALSE
		oParser = new parserFormula('ISODD(-(2^31-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(-(2^31-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Largest negative odd 32-bit integer, returns FALSE');
		// Case #5: Formula. Largest Excel number (truncated to integer), returns FALSE
		oParser = new parserFormula('ISODD(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISODD(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Largest Excel number (truncated to integer), returns FALSE');

		// Need to fix: boolean handle
		// Case #7: Boolean. Boolean input, returns #VALUE!
		// Case #8: Boolean. Boolean input, returns #VALUE!


		testArrayFormula2(assert, "ISODD", 1, 1, true, null);
	});

    QUnit.test("Test: ISREF", function (assert) {
		oParser = new parserFormula("ISREF(G0)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Number. Number input, returns TRUE
		oParser = new parserFormula('ISREF(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Number input, returns TRUE');
		// Case #2: Number. Zero, returns TRUE
		oParser = new parserFormula('ISREF(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Zero, returns TRUE');
		// Case #3: Empty. Empty argument (blank cell), returns TRUE
		oParser = new parserFormula('ISREF("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Empty. Empty argument (blank cell), returns TRUE');
		// Case #4: Error. Error (#DIV/0!), returns TRUE
		oParser = new parserFormula('ISREF(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. Error (#DIV/0!), returns TRUE');
		// Case #5: Error. Error (#VALUE!), returns TRUE
		oParser = new parserFormula('ISREF(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. Error (#VALUE!), returns TRUE');
		// Case #6: Error. Error (#N/A), returns TRUE
		oParser = new parserFormula('ISREF(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. Error (#N/A), returns TRUE');
		// Case #7: Boolean. Boolean TRUE, returns TRUE
		oParser = new parserFormula('ISREF(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean. Boolean TRUE, returns TRUE');
		// Case #8: Boolean. Boolean FALSE, returns TRUE
		oParser = new parserFormula('ISREF(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean. Boolean FALSE, returns TRUE');
		// Case #9: Formula. Date formula (serial number), returns TRUE
		oParser = new parserFormula('ISREF(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Date formula (serial number), returns TRUE');
		// Case #10: Formula. Time formula (serial number), returns TRUE
		oParser = new parserFormula('ISREF(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Time formula (serial number), returns TRUE');
		// Case #11: Reference link. Reference to empty cell, returns TRUE
		oParser = new parserFormula('ISREF(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to empty cell, returns TRUE');
		// Case #12: Reference link. Reference to cell with number, returns TRUE
		oParser = new parserFormula('ISREF(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with number, returns TRUE');
		// Case #13: Area. Single-cell range with number, returns TRUE
		oParser = new parserFormula('ISREF(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Single-cell range with number, returns TRUE');
		// Case #14: Area. Multi-cell range with number, returns TRUE
		oParser = new parserFormula('ISREF(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Multi-cell range with number, returns TRUE');
		// Case #15: Array. Array with number, returns TRUE
		oParser = new parserFormula('ISREF({123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF({123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with number, returns TRUE');
		// Case #16: Name. Named range with error, returns TRUE
		oParser = new parserFormula('ISREF(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name. Named range with error, returns TRUE');
		// Case #17: Name3D. 3D named range with boolean, returns TRUE
		oParser = new parserFormula('ISREF(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D. 3D named range with boolean, returns TRUE');
		// Case #18: Ref3D. 3D reference to cell with number, returns TRUE
		oParser = new parserFormula('ISREF(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Ref3D. 3D reference to cell with number, returns TRUE');
		// Case #19: Area3D. 3D single-cell range with error, returns TRUE
		oParser = new parserFormula('ISREF(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area3D. 3D single-cell range with error, returns TRUE');
		// Case #20: Table. Table column with empty cell, returns TRUE
		oParser = new parserFormula('ISREF(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Table. Table column with empty cell, returns TRUE');
		// Case #21: Formula. Nested IF returning number, returns TRUE
		oParser = new parserFormula('ISREF(IF(TRUE,123,"text"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(IF(TRUE,123,"text")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested IF returning number, returns TRUE');
		// Case #22: Formula. Nested SQRT formula returning number, returns TRUE
		oParser = new parserFormula('ISREF(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested SQRT formula returning number, returns TRUE');
		// Case #23: Time. Time (0.0, number), returns TRUE
		oParser = new parserFormula('ISREF(TIME(0,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TIME(0,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Time. Time (0.0, number), returns TRUE');
		// Case #24: Formula. Volatile NOW formula (serial number), returns TRUE
		oParser = new parserFormula('ISREF(NOW())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(NOW()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Volatile NOW formula (serial number), returns TRUE');
		// Case #25: Reference link. Reference to cell with boolean, returns TRUE
		oParser = new parserFormula('ISREF(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with boolean, returns TRUE');

		// Negative cases:
		// Case #1: String. Text string, returns FALSE
		oParser = new parserFormula('ISREF("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Text string, returns FALSE');
		// Case #2: String. Empty string, returns FALSE
		oParser = new parserFormula('ISREF("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Empty string, returns FALSE');
		// Case #3: String. Numeric string (still text), returns FALSE
		oParser = new parserFormula('ISREF("123")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF("123") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Numeric string (still text), returns FALSE');
		// Case #4: String. String with space, returns FALSE
		oParser = new parserFormula('ISREF(" ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(" ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. String with space, returns FALSE');
		// Case #5: Formula. Nested TEXT formula returning string, returns FALSE
		oParser = new parserFormula('ISREF(TEXT(123,"0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TEXT(123,"0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested TEXT formula returning string, returns FALSE');
		// Case #6: Reference link. Reference to cell with text, returns FALSE
		oParser = new parserFormula('ISREF(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with text, returns FALSE');
		// Case #7: Reference link. Reference to cell with empty string, returns FALSE
		oParser = new parserFormula('ISREF(A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with empty string, returns FALSE');
		// Case #8: Area. Single-cell range with text, returns FALSE
		oParser = new parserFormula('ISREF(A108:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A108:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Single-cell range with text, returns FALSE');
		// Case #9: Area. Multi-cell range with text, returns FALSE
		oParser = new parserFormula('ISREF(A109:A110)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A109:A110) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Multi-cell range with text, returns FALSE');
		// Case #10: Array. Array with text, returns FALSE
		oParser = new parserFormula('ISREF({"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF({"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with text, returns FALSE');
		// Case #11: Array. Array with text as first element, returns FALSE
		oParser = new parserFormula('ISREF({"abc",123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF({"abc",123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with text as first element, returns FALSE');
		// Case #12: Name. Named range with text, returns FALSE
		oParser = new parserFormula('ISREF(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name. Named range with text, returns FALSE');
		// Case #13: Name. Named range with empty string, returns FALSE
		oParser = new parserFormula('ISREF(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name. Named range with empty string, returns FALSE');
		// Case #14: Name3D. 3D named range with text, returns FALSE
		oParser = new parserFormula('ISREF(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name3D. 3D named range with text, returns FALSE');
		// Case #15: Ref3D. 3D reference to cell with text, returns FALSE
		oParser = new parserFormula('ISREF(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Ref3D. 3D reference to cell with text, returns FALSE');
		// Case #16: Area3D. 3D multi-cell range with text, returns FALSE
		oParser = new parserFormula('ISREF(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area3D. 3D multi-cell range with text, returns FALSE');
		// Case #17: Table. Table column with text, returns FALSE
		oParser = new parserFormula('ISREF(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Table. Table column with text, returns FALSE');
		// Case #18: Formula. Nested IF returning text, returns FALSE
		oParser = new parserFormula('ISREF(IF(TRUE,"text",123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(IF(TRUE,"text",123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested IF returning text, returns FALSE');
		// Case #19: Formula. Nested CONCAT formula returning text, returns FALSE
		oParser = new parserFormula('ISREF(CONCAT("a","b"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(CONCAT("a","b")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested CONCAT formula returning text, returns FALSE');
		// Case #20: Reference link. Reference to cell with numeric string, returns FALSE
		oParser = new parserFormula('ISREF(A111)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(A111) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with numeric string, returns FALSE');

		// Bounded cases:
		// Case #1: Number. Smallest positive Excel number, returns TRUE
		oParser = new parserFormula('ISREF(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Smallest positive Excel number, returns TRUE');
		// Case #2: Number. Largest Excel number, returns TRUE
		oParser = new parserFormula('ISREF(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Largest Excel number, returns TRUE');
		// Case #3: Date. Minimum Excel date (serial number 1), returns TRUE
		oParser = new parserFormula('ISREF(DATE(1900,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(DATE(1900,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Date. Minimum Excel date (serial number 1), returns TRUE');
		// Case #4: Date. Maximum Excel date (serial number 2958465), returns TRUE
		oParser = new parserFormula('ISREF(DATE(9999,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(DATE(9999,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Date. Maximum Excel date (serial number 2958465), returns TRUE');
		// Case #5: Formula. Formula resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISREF(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISREF(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Formula resulting in #DIV/0!, returns TRUE');


		testArrayFormula2(assert, "ISREF", 1, 1, null, true);
	});

	QUnit.test("Test: ISTEXT", function (assert) {
		ws.getRange2("S7").setValue("test");

		oParser = new parserFormula("ISTEXT(S7)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("Text2");
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
		// Case #1: Number. Number input, returns TRUE
		oParser = new parserFormula('ISTEXT(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Number input, returns TRUE');
		// Case #2: Number. Zero, returns TRUE
		oParser = new parserFormula('ISTEXT(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Zero, returns TRUE');
		// Case #3: Empty. Empty argument (blank cell), returns TRUE
		oParser = new parserFormula('ISTEXT("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Empty. Empty argument (blank cell), returns TRUE');
		// Case #4: Error. Error (#DIV/0!), returns TRUE
		oParser = new parserFormula('ISTEXT(#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. Error (#DIV/0!), returns TRUE');
		// Case #5: Error. Error (#VALUE!), returns TRUE
		oParser = new parserFormula('ISTEXT(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. Error (#VALUE!), returns TRUE');
		// Case #6: Error. Error (#N/A), returns TRUE
		oParser = new parserFormula('ISTEXT(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Error. Error (#N/A), returns TRUE');
		// Case #7: Boolean. Boolean TRUE, returns TRUE
		oParser = new parserFormula('ISTEXT(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean. Boolean TRUE, returns TRUE');
		// Case #8: Boolean. Boolean FALSE, returns TRUE
		oParser = new parserFormula('ISTEXT(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean. Boolean FALSE, returns TRUE');
		// Case #9: Formula. Date formula (serial number), returns TRUE
		oParser = new parserFormula('ISTEXT(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Date formula (serial number), returns TRUE');
		// Case #10: Formula. Time formula (serial number), returns TRUE
		oParser = new parserFormula('ISTEXT(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Time formula (serial number), returns TRUE');
		// Case #11: Reference link. Reference to empty cell, returns TRUE
		oParser = new parserFormula('ISTEXT(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to empty cell, returns TRUE');
		// Case #12: Reference link. Reference to cell with number, returns TRUE
		oParser = new parserFormula('ISTEXT(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with number, returns TRUE');
		// Case #13: Area. Single-cell range with number, returns TRUE
		oParser = new parserFormula('ISTEXT(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Single-cell range with number, returns TRUE');
		// Case #14: Area. Multi-cell range with number, returns TRUE
		oParser = new parserFormula('ISTEXT(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Multi-cell range with number, returns TRUE');
		// Case #15: Array. Array with number, returns TRUE
		oParser = new parserFormula('ISTEXT({123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT({123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with number, returns TRUE');
		// Case #16: Name. Named range with error, returns TRUE
		oParser = new parserFormula('ISTEXT(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with error, returns TRUE');
		// Case #17: Name3D. 3D named range with boolean, returns TRUE
		oParser = new parserFormula('ISTEXT(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with boolean, returns TRUE');
		// Case #18: Ref3D. 3D reference to cell with number, returns TRUE
		oParser = new parserFormula('ISTEXT(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with number, returns TRUE');
		// Case #19: Area3D. 3D single-cell range with error, returns TRUE
		oParser = new parserFormula('ISTEXT(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with error, returns TRUE');
		// Case #20: Table. Table column with empty cell, returns TRUE
		oParser = new parserFormula('ISTEXT(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with empty cell, returns TRUE');
		// Case #21: Formula. Nested IF returning number, returns TRUE
		oParser = new parserFormula('ISTEXT(IF(TRUE,123,"text"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(IF(TRUE,123,"text")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested IF returning number, returns TRUE');
		// Case #22: Formula. Nested SQRT formula returning number, returns TRUE
		oParser = new parserFormula('ISTEXT(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested SQRT formula returning number, returns TRUE');
		// Case #23: Time. Time (0.0, number), returns TRUE
		oParser = new parserFormula('ISTEXT(TIME(0,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TIME(0,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Time. Time (0.0, number), returns TRUE');
		// Case #24: Formula. Volatile NOW formula (serial number), returns TRUE
		oParser = new parserFormula('ISTEXT(NOW())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(NOW()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Volatile NOW formula (serial number), returns TRUE');
		// Case #25: Reference link. Reference to cell with boolean, returns TRUE
		oParser = new parserFormula('ISTEXT(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with boolean, returns TRUE');

		// Negative cases:
		// Case #1: String. Text string, returns FALSE
		oParser = new parserFormula('ISTEXT("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String. Text string, returns FALSE');
		// Case #2: String. Empty string, returns FALSE
		oParser = new parserFormula('ISTEXT("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String. Empty string, returns FALSE');
		// Case #3: String. Numeric string (still text), returns FALSE
		oParser = new parserFormula('ISTEXT("123")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT("123") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String. Numeric string (still text), returns FALSE');
		// Case #4: String. String with space, returns FALSE
		oParser = new parserFormula('ISTEXT(" ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(" ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String. String with space, returns FALSE');
		// Case #5: Formula. Nested TEXT formula returning string, returns FALSE
		oParser = new parserFormula('ISTEXT(TEXT(123,"0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TEXT(123,"0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Formula. Nested TEXT formula returning string, returns FALSE');
		// Case #6: Reference link. Reference to cell with text, returns FALSE
		oParser = new parserFormula('ISTEXT(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with text, returns FALSE');
		// Case #7: Reference link. Reference to cell with empty string, returns FALSE
		oParser = new parserFormula('ISTEXT(A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with empty string, returns FALSE');
		// Case #8: Area. Single-cell range with text, returns FALSE
		oParser = new parserFormula('ISTEXT(A108:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A108:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Single-cell range with text, returns FALSE');
		// Case #9: Area. Multi-cell range with text, returns FALSE
		oParser = new parserFormula('ISTEXT(A109:A110)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A109:A110) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range with text, returns FALSE');
		// Case #10: Array. Array with text, returns FALSE
		oParser = new parserFormula('ISTEXT({"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT({"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Array. Array with text, returns FALSE');
		// Case #11: Array. Array with text as first element, returns FALSE
		oParser = new parserFormula('ISTEXT({"abc",123})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT({"abc",123}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Array. Array with text as first element, returns FALSE');
		// Case #12: Name. Named range with text, returns FALSE
		oParser = new parserFormula('ISTEXT(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with text, returns FALSE');
		// Case #13: Name. Named range with empty string, returns FALSE
		oParser = new parserFormula('ISTEXT(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with empty string, returns FALSE');
		// Case #14: Name3D. 3D named range with text, returns FALSE
		oParser = new parserFormula('ISTEXT(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with text, returns FALSE');
		// Case #15: Ref3D. 3D reference to cell with text, returns FALSE
		oParser = new parserFormula('ISTEXT(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Ref3D. 3D reference to cell with text, returns FALSE');
		// Case #16: Area3D. 3D multi-cell range with text, returns FALSE
		oParser = new parserFormula('ISTEXT(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range with text, returns FALSE');
		// Case #17: Table. Table column with text, returns FALSE
		oParser = new parserFormula('ISTEXT(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Table. Table column with text, returns FALSE');
		// Case #18: Formula. Nested IF returning text, returns FALSE
		oParser = new parserFormula('ISTEXT(IF(TRUE,"text",123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(IF(TRUE,"text",123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Formula. Nested IF returning text, returns FALSE');
		// Case #19: Formula. Nested CONCAT formula returning text, returns FALSE
		oParser = new parserFormula('ISTEXT(CONCAT("a","b"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(CONCAT("a","b")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Formula. Nested CONCAT formula returning text, returns FALSE');
		// Case #20: Reference link. Reference to cell with numeric string, returns FALSE
		oParser = new parserFormula('ISTEXT(A111)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(A111) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with numeric string, returns FALSE');

		// Bounded cases:
		// Case #1: Number. Smallest positive Excel number, returns TRUE
		oParser = new parserFormula('ISTEXT(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Smallest positive Excel number, returns TRUE');
		// Case #2: Number. Largest Excel number, returns TRUE
		oParser = new parserFormula('ISTEXT(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Largest Excel number, returns TRUE');
		// Case #3: Date. Minimum Excel date (serial number 1), returns TRUE
		oParser = new parserFormula('ISTEXT(DATE(1900,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(DATE(1900,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Date. Minimum Excel date (serial number 1), returns TRUE');
		// Case #4: Date. Maximum Excel date (serial number 2958465), returns TRUE
		oParser = new parserFormula('ISTEXT(DATE(9999,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(DATE(9999,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Date. Maximum Excel date (serial number 2958465), returns TRUE');
		// Case #5: Formula. Formula resulting in #DIV/0!, returns TRUE
		oParser = new parserFormula('ISTEXT(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISTEXT(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Formula. Formula resulting in #DIV/0!, returns TRUE');

		// Need to fix:
		// Case #14: Area. Multi-cell range with number, returns TRUE

		testArrayFormula2(assert, "ISTEXT", 1, 1);
	});

	QUnit.test("Test: \"N\"", function (assert) {

		ws.getRange2("A2").setValue("7");
		ws.getRange2("A3").setValue("Even");
		ws.getRange2("A4").setValue("TRUE");
		ws.getRange2("A5").setValue("4/17/2011");

		oParser = new parserFormula("N(A2)", "A7", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 7);

		oParser = new parserFormula("N(A3)", "A7", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula("N(A4)", "A7", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("N(A5)", "A7", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 40650);

		oParser = new parserFormula('N("7")', "A7", ws);
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
		// Case #0: Number. Basic valid input: integer. 1 argument used.
		oParser = new parserFormula('N(15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Number. Basic valid input: integer. 1 argument used.');
		// Case #1: Number. Basic valid input: float number. 1 argument used.
		oParser = new parserFormula('N(3.7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(3.7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3.7, 'Test: Positive case: Number. Basic valid input: float number. 1 argument used.');
		// Case #2: String. String convertible to number. 1 argument used.
		oParser = new parserFormula('N("25")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N("25") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: String. String convertible to number. 1 argument used.');
		// Case #3: Formula. Nested formula returning number (3). 1 argument used.
		oParser = new parserFormula('N(SQRT(9))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(SQRT(9)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Formula. Nested formula returning number (3). 1 argument used.');
		// Case #4: Reference link. Reference to cell with valid number (3). 1 argument used.
		oParser = new parserFormula('N(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to cell with valid number (3). 1 argument used.');
		// Case #5: Area. Single-cell range with valid number (1.1). 1 argument used.
		oParser = new parserFormula('N(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area. Single-cell range with valid number (1.1). 1 argument used.');
		// Case #6: Array. Array with single numeric element. 1 argument used.
		oParser = new parserFormula('N({5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N({5}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 5, 'Test: Positive case: Array. Array with single numeric element. 1 argument used.');
		// Case #7: Name. Named range with valid number (3). 1 argument used.
		oParser = new parserFormula('N(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -0.5, 'Test: Positive case: Name. Named range with valid number (3). 1 argument used.');
		// Case #8: Name3D. 3D named range with valid number (2). 1 argument used.
		oParser = new parserFormula('N(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -0.5, 'Test: Positive case: Name3D. 3D named range with valid number (2). 1 argument used.');
		// Case #9: Ref3D. 3D reference to cell with valid number (2). 1 argument used.
		oParser = new parserFormula('N(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. 3D reference to cell with valid number (2). 1 argument used.');
		// Case #10: Area3D. 3D single-cell range with valid number (2). 1 argument used.
		oParser = new parserFormula('N(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area3D. 3D single-cell range with valid number (2). 1 argument used.');
		// Case #11: Table. Table structured reference with valid number (3). 1 argument used.
		oParser = new parserFormula('N(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table structured reference with valid number (3). 1 argument used.');
		// Case #12: Date. Date as serial number (45808). 1 argument used.
		oParser = new parserFormula('N(DATE(2025,6,30))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(DATE(2025,6,30)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45838, 'Test: Positive case: Date. Date as serial number (45808). 1 argument used.');
		// Case #13: Time. Time as fraction of day (0.60416667). 1 argument used.
		oParser = new parserFormula('N(TIME(14,30,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(TIME(14,30,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(3), "0.604", 'Test: Positive case: Time. Time as fraction of day (0.60416667). 1 argument used.');
		// Case #14: Formula. Nested ABS formula returning number (8). 1 argument used.
		oParser = new parserFormula('N(ABS(-8))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(ABS(-8)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Formula. Nested ABS formula returning number (8). 1 argument used.');
		// Case #15: Boolean. Boolean TRUE converts to 1. 1 argument used.
		oParser = new parserFormula('N(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Boolean. Boolean TRUE converts to 1. 1 argument used.');
		// Case #16: Boolean. Boolean FALSE converts to 0. 1 argument used.
		oParser = new parserFormula('N(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Boolean. Boolean FALSE converts to 0. 1 argument used.');
		// Case #17: String. Short date string convertible to number (6). 1 argument used.
		oParser = new parserFormula('N("6/6")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N("6/6") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: String. Short date string convertible to number (6). 1 argument used.');
		// Case #18: Array. Multi-element array, returns first element (12). 1 argument used.
		oParser = new parserFormula('N({12,24})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N({12,24}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 12, 'Test: Positive case: Array. Multi-element array, returns first element (12). 1 argument used.');
		// Case #19: Formula. Nested IF returning valid number (10). 1 argument used.
		oParser = new parserFormula('N(IF(TRUE,10,4))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(IF(TRUE,10,4)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Formula. Nested IF returning valid number (10). 1 argument used.');
		// Case #20: Formula. Nested ROUND formula returning number (7.7). 1 argument used.
		oParser = new parserFormula('N(ROUND(7.666,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(ROUND(7.666,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7.7, 'Test: Positive case: Formula. Nested ROUND formula returning number (7.7). 1 argument used.');
		// Case #21: Formula. N with nested SUM formula returning number (7). 1 argument used.
		oParser = new parserFormula('N(SUM(5,2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(SUM(5,2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Formula. N with nested SUM formula returning number (7). 1 argument used.');
		// Case #22: String. String in scientific notation convertible to number (1000000000000). 1 argument used.
		oParser = new parserFormula('N("1E+12")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N("1E+12") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: String. String in scientific notation convertible to number (1000000000000). 1 argument used.');

		// Negative cases:
		// Case #1: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('N(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #2: String. Non-numeric string returns 0. 1 argument used.
		oParser = new parserFormula('N("xyz")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N("xyz") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Non-numeric string returns 0. 1 argument used.');
		// Case #3: Empty. Reference to empty cell returns 0. 1 argument used.
		oParser = new parserFormula('N(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Negative case: Empty. Reference to empty cell returns 0. 1 argument used.');
		// Case #4: Area. Multi-cell range returns first cell’s value (1.1). 1 argument used.
		oParser = new parserFormula('N(A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(A101:A102) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area. Multi-cell range returns first cell’s value (1.1). 1 argument used.');
		// Case #5: Ref3D. 3D reference to cell with text (abc) returns 0. 1 argument used.
		oParser = new parserFormula('N(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: Ref3D. 3D reference to cell with text (abc) returns 0. 1 argument used.');
		// Case #6: Name. Named range with text (invalid) returns 0. 1 argument used.
		oParser = new parserFormula('N(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Test: Negative case: Name. Named range with text (invalid) returns 0. 1 argument used.');
		// Case #7: Name3D. 3D named range. 1 argument used.
		oParser = new parserFormula('N(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.8, 'Test: Negative case: Name3D. 3D named range. 1 argument used.');
		// Case #9: Formula. Formula resulting in #DIV/0! propagates error. 1 argument used.
		oParser = new parserFormula('N(DIVIDE(5,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(DIVIDE(5,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Formula. Formula resulting in #DIV/0! propagates error. 1 argument used.');
		// Case #10: String. Empty string returns 0. 1 argument used.
		oParser = new parserFormula('N("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Empty string returns 0. 1 argument used.');
		// Case #11: Array. Array with booleans, returns first element (1). 1 argument used.
		oParser = new parserFormula('N({TRUE,FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N({TRUE,FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Negative case: Array. Array with booleans, returns first element (1). 1 argument used.');
		// Case #12: Area3D. 3D multi-cell range, returns first cell’s value (2). 1 argument used.
		oParser = new parserFormula('N(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(Sheet2!A1:A2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Negative case: Area3D. 3D multi-cell range, returns first cell’s value (2). 1 argument used.');
		// Case #13: Formula. Formula resulting in #NUM! propagates error. 1 argument used.
		oParser = new parserFormula('N(SQRT(-4))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(SQRT(-4)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. 1 argument used.');
		// Case #14: String. Text string returns 0. 1 argument used.
		oParser = new parserFormula('N("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Text string returns 0. 1 argument used.');
		// Case #15: Array. Array with non-numeric string returns 0. 1 argument used.
		oParser = new parserFormula('N({"xyz"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N({"xyz"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 0, 'Test: Negative case: Array. Array with non-numeric string returns 0. 1 argument used.');
		// Case #16: Reference link. Reference to cell with non-numeric value (0.1) returns 0.1. 1 argument used.
		oParser = new parserFormula('N(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Reference link. Reference to cell with non-numeric value (0.1) returns 0.1. 1 argument used.');
		// Case #17: Formula. Formula resulting in #N/A propagates error. 1 argument used.
		oParser = new parserFormula('N(VLOOKUP("z",A1:B1,2,FALSE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(VLOOKUP("z",A1:B1,2,FALSE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula. Formula resulting in #N/A propagates error. 1 argument used.');
		// Case #18: String. Non-numeric string with negative sign returns 0. 1 argument used.
		oParser = new parserFormula('N("-text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N("-text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Non-numeric string with negative sign returns 0. 1 argument used.');
		// Case #19: Time. Invalid time returns 0. 1 argument used.
		oParser = new parserFormula('N(TIME(26,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(TIME(26,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(2), "0.08", 'Test: Negative case: Time. Invalid time returns 0. 1 argument used.');
		// Case #20: Date. Date before valid range (0) returns 0. 1 argument used.
		oParser = new parserFormula('N(DATE(1899,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(DATE(1899,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 693962, 'Test: Negative case: Date. Date before valid range (0) returns 0. 1 argument used.');

		// Bounded cases:
		// Case #1: Number. Smallest valid positive number. 1 argument used.
		oParser = new parserFormula('N(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1e-307, 'Test: Bounded case: Number. Smallest valid positive number. 1 argument used.');
		// Case #2: Number. Largest valid Excel number. 1 argument used.
		oParser = new parserFormula('N(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9.99999999999999e+307, 'Test: Bounded case: Number. Largest valid Excel number. 1 argument used.');
		// Case #3: Date. Smallest valid date serial number (1900-01-01). 1 argument used.
		oParser = new parserFormula('N(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Date. Smallest valid date serial number (1900-01-01). 1 argument used.');
		// Case #4: Date. Largest valid date serial number (9999-12-31). 1 argument used.
		oParser = new parserFormula('N(2958465)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: N(2958465) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958465, 'Test: Bounded case: Date. Largest valid date serial number (9999-12-31). 1 argument used.');

		// Need to fix: area/ area3D handle
		// Case #4: Area. Multi-cell range returns first cell’s value (1.1). 1 argument used.
		// Case #12: Area3D. 3D multi-cell range, returns first cell’s value (2). 1 argument used.

		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "N", 1, 1);
	});

	QUnit.test("Test: \"SHEET\"", function (assert) {

		oParser = new parserFormula("SHEET(Hi_Temps)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?");

		ws.getRange2("A100:C214").cleanAll();
		// Data for reference link. Use A100-A115
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("A102").setValue("Text");
		ws.getRange2("A103").setValue("TRUE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("5"); // Num (Column1)
		ws.getRange2("B601").setValue("10"); // Num (Column2)
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

		wb.createWorksheet(0, "Sheet3");
		wb.createWorksheet(0, "Sheet4");
		wb.createWorksheet(0, "Sheet5");

		// Positive cases:
		// Case #1: Empty. No argument provided, returns sheet number of the cell containing the formula. 0 arguments used.
		oParser = new parserFormula('SHEET()', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET() is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Empty. No argument provided, returns sheet number of the cell containing the formula. 0 arguments used.');
		// Case #2: Reference link. Reference to a single cell. Returns sheet number of the referenced cell. 1 argument used.
		oParser = new parserFormula('SHEET(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Reference link. Reference to a single cell. Returns sheet number of the referenced cell. 1 argument used.');
		// Case #3: Area. Single-cell range. Returns sheet number of the first cell in the range. 1 argument used.
		oParser = new parserFormula('SHEET(A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(A101:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area. Single-cell range. Returns sheet number of the first cell in the range. 1 argument used.');
		// Case #4: Name. Named range referring to a single cell or range. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Name. Named range referring to a single cell or range. Returns sheet number. 1 argument used.');
		// Case #5: Name3D. 3D named range referring to a cell or range. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Name3D. 3D named range referring to a cell or range. Returns sheet number. 1 argument used.');
		// Case #6: Ref3D. 3D reference to a single cell. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Ref3D. 3D reference to a single cell. Returns sheet number. 1 argument used.');
		// Case #7: Area3D. 3D single-cell range. Returns sheet number of the first cell. 1 argument used.
		oParser = new parserFormula('SHEET(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Area3D. 3D single-cell range. Returns sheet number of the first cell. 1 argument used.');
		// Case #8: Table. Table structured reference. Returns sheet number of the first cell in the column. 1 argument used.
		oParser = new parserFormula('SHEET(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Table. Table structured reference. Returns sheet number of the first cell in the column. 1 argument used.');
		// Case #9: Formula. Formula resolving to a single cell reference. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(IF(TRUE,A100,B100))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(IF(TRUE,A100,B100)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Formula. Formula resolving to a single cell reference. Returns sheet number. 1 argument used.');
		// Case #10: String. Valid sheet name as string. Returns sheet number of the specified sheet. 1 argument used.
		oParser = new parserFormula('SHEET("Sheet2")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET("Sheet2") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: String. Valid sheet name as string. Returns sheet number of the specified sheet. 1 argument used.');
		// Case #11: Formula. Formula resolving to a valid cell reference. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(INDIRECT("Sheet2!A1"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(INDIRECT("Sheet2!A1")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Formula. Formula resolving to a valid cell reference. Returns sheet number. 1 argument used.');
		// Case #12: Reference link. Reference to a cell with numeric value. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Reference link. Reference to a cell with numeric value. Returns sheet number. 1 argument used.');
		// Case #13: Area. Single-cell range explicitly defined. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area. Single-cell range explicitly defined. Returns sheet number. 1 argument used.');
		// Case #14: Name. Named range referring to a range. Returns sheet number of first cell. 1 argument used.
		oParser = new parserFormula('SHEET(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Name. Named range referring to a range. Returns sheet number of first cell. 1 argument used.');
		// Case #15: Name3D. 3D named range referring to a range. Returns sheet number of first cell. 1 argument used.
		oParser = new parserFormula('SHEET(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Name3D. 3D named range referring to a range. Returns sheet number of first cell. 1 argument used.');
		// Case #16: Ref3D. 3D reference to a single cell. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Ref3D. 3D reference to a single cell. Returns sheet number. 1 argument used.');
		// Case #17: Formula. SHEET inside SUM formula, returns sum of sheet numbers for range. 1 argument used.
		oParser = new parserFormula('SUM(SHEET(A1:A3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(SHEET(A1:A3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Formula. SHEET inside SUM formula, returns sum of sheet numbers for range. 1 argument used.');
		// Case #18: Reference link. Reference to a cell with string value. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Reference link. Reference to a cell with string value. Returns sheet number. 1 argument used.');
		// Case #19: Area. Vertical range, returns sheet number of first cell. 1 argument used.
		oParser = new parserFormula('SHEET(A1:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(A1:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area. Vertical range, returns sheet number of first cell. 1 argument used.');
		// Case #20: Table. Table structured reference with text values. Returns sheet number of first cell. 1 argument used.
		oParser = new parserFormula('SHEET(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Table. Table structured reference with text values. Returns sheet number of first cell. 1 argument used.');
		// Case #21: String. Valid sheet name as string (alternative sheet name format). Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET("3D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET("3D") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Test: Positive case: String. Valid sheet name as string (alternative sheet name format). Returns sheet number. 1 argument used.');
		// Case #22: Formula. Formula resolving to a valid sheet reference. Returns sheet number. 1 argument used.
		oParser = new parserFormula('SHEET(ADDRESS(1,1,4,,"Sheet2"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(ADDRESS(1,1,4,,"Sheet2")) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Formula. Formula resolving to a valid sheet reference. Returns sheet number. 1 argument used.');

		// Negative cases:
		// Case #1: String. Invalid sheet name. Returns #N/A. 1 argument used.
		oParser = new parserFormula('SHEET("badSheetName")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET("badSheetName") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String. Invalid sheet name. Returns #N/A. 1 argument used.');
		// Case #2: Number. Numeric value not a reference or sheet name. Returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SHEET(5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number. Numeric value not a reference or sheet name. Returns #VALUE!. 1 argument used.');
		// Case #3: Error. Error input propagates #DIV/0. 1 argument used.
		oParser = new parserFormula('SHEET(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Error. Error input propagates #N/A. 1 argument used.');
		// Case #5: Name. Named range with text value. Returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SHEET(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Negative case: Name. Named range with text value. Returns #VALUE!. 1 argument used.');
		// Case #6: Name3D. 3D named range with text value. Returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SHEET(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Negative case: Name3D. 3D named range with text value. Returns #VALUE!. 1 argument used.');
		// Case #7: Ref3D. 3D reference to non-existent cell. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEET(Sheet2!Z1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Sheet2!Z1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Negative case: Ref3D. 3D reference to non-existent cell. Returns #REF!. 1 argument used.');
		// Case #9: Array. Array constant not a reference or sheet name. Returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SHEET({1;2;3})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET({1;2;3}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Array. Array constant not a reference or sheet name. Returns #VALUE!. 1 argument used.');
		// Case #10: Formula. Formula resulting in #NUM! propagates error. Returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SHEET(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. Returns #VALUE!. 1 argument used.');
		// Case #11: Boolean. Boolean value not a reference or sheet name. Returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SHEET(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Boolean. Boolean value not a reference or sheet name. Returns #VALUE!. 1 argument used.');
		// Case #13: String. Date string not a valid sheet name. Returns #N/A. 1 argument used.
		oParser = new parserFormula('SHEET("01.01.2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET("01.01.2025") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String. Date string not a valid sheet name. Returns #N/A. 1 argument used.');
		// Case #14: Formula. Formula resolving to invalid date (#NUM!). Returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SHEET(DATE(2025,13,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(DATE(2025,13,1)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula. Formula resolving to invalid date (#NUM!). Returns #VALUE!. 1 argument used.');
		// Case #17: Table. Table column with invalid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEET(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Negative case: Table. Table column with invalid reference. Returns #REF!. 1 argument used.');
		// Case #18: Formula. Formula resolving to invalid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEET(INDIRECT("Z1"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(INDIRECT("Z1")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Negative case: Formula. Formula resolving to invalid reference. Returns #REF!. 1 argument used.');
		// Case #20: Ref3D. 3D reference to invalid cell. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEET(Sheet2!#REF)', 'A2', ws);
		assert.ok(oParser.parse() === false, 'Test: SHEET(Sheet2!#REF) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Ref3D. 3D reference to invalid cell. Returns #REF!. 1 argument used.');

		// Bounded cases:
		// Case #1: Reference link. Reference to cell in first sheet (Sheet1). Returns 1. 1 argument used.
		oParser = new parserFormula('SHEET(A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Bounded case: Reference link. Reference to cell in first sheet (Sheet1). Returns 1. 1 argument used.');
		// Case #2: Ref3D. Reference to cell in second sheet (Sheet2/3D). Returns 2. 1 argument used.
		oParser = new parserFormula('SHEET(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Bounded case: Ref3D. Reference to cell in second sheet (Sheet2/3D). Returns 2. 1 argument used.');
		// Case #3: Area. Single-cell range in first sheet. Returns 1. 1 argument used.
		oParser = new parserFormula('SHEET(A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Bounded case: Area. Single-cell range in first sheet. Returns 1. 1 argument used.');
		// Case #4: Area3D. Single-cell 3D range in second sheet. Returns 2. 1 argument used.
		oParser = new parserFormula('SHEET(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEET(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Bounded case: Area3D. Single-cell 3D range in second sheet. Returns 2. 1 argument used.');

		// Delete created worksheets for tests
		wb.removeWorksheet(0);
		wb.removeWorksheet(0);
		wb.removeWorksheet(0);

		// Need to fix: different error types(should be #N/A instead #VALUE)
		// Case #21: String. Valid sheet name as string (alternative sheet name format). Returns sheet number. 1 argument used.
		// Case #22: Formula. Formula resolving to a valid sheet reference. Returns sheet number. 1 argument used.
		// Case #1: String. Invalid sheet name. Returns #N/A. 1 argument used.
		// Case #2: Number. Numeric value not a reference or sheet name. Returns #VALUE!. 1 argument used.
		// Case #9: Array. Array constant not a reference or sheet name. Returns #VALUE!. 1 argument used.
		// Case #11: Boolean. Boolean value not a reference or sheet name. Returns #VALUE!. 1 argument used.
		// Case #13: String. Date string not a valid sheet name. Returns #N/A. 1 argument used.
		// Case #14: Formula. Formula resolving to invalid date (#NUM!). Returns #VALUE!. 1 argument used.


		testArrayFormula2(assert, "SHEET", 1, 1, null, true);
	});

	QUnit.test("Test: \"SHEETS\"", function (assert) {

		oParser = new parserFormula("SHEETS(Hi_Temps)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?");

		oParser = new parserFormula("SHEETS()", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

        ws.getRange2("A100:C214").cleanAll();
		// Data for reference link. Use A100-A115
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("A102").setValue("Text");
		ws.getRange2("A103").setValue("TRUE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("5"); // Num (Column1)
		ws.getRange2("B601").setValue("10"); // Num (Column2)
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

		wb.createWorksheet(0, "Sheet3");
		wb.createWorksheet(0, "Sheet4");
		wb.createWorksheet(0, "Sheet5");


		// Positive cases:
		// Case #1: Empty. No argument provided, returns number of sheets in the workbook. 0 arguments used.
		oParser = new parserFormula('SHEETS()', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS() is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Empty. No argument provided, returns number of sheets in the workbook. 0 arguments used.');
		// Case #2: Reference link. Reference to a single cell. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to a single cell. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #3: Area. Single-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A101:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area. Single-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #4: Name. Named range referring to a cell or range. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name. Named range referring to a cell or range. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #5: Name3D. 3D named range referring to a cell or range. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name3D. 3D named range referring to a cell or range. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #6: Ref3D. 3D reference to a single cell. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. 3D reference to a single cell. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #7: Area3D. 3D single-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area3D. 3D single-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #8: Table. Table structured reference. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table structured reference. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #9: Formula. Formula resolving to a single cell reference. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(IF(TRUE,A100,B100))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(IF(TRUE,A100,B100)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Formula resolving to a single cell reference. Returns number of sheets (1). 1 argument used.');
		// Case #10: Reference link. Reference to a cell with numeric value. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to a cell with numeric value. Returns number of sheets (1). 1 argument used.');
		// Case #11: Area. Single-cell range explicitly defined. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area. Single-cell range explicitly defined. Returns number of sheets (1). 1 argument used.');
		// Case #12: Formula. Formula resolving to a valid cell reference. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(INDIRECT("Sheet2!A1"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(INDIRECT("Sheet2!A1")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Formula resolving to a valid cell reference. Returns number of sheets (1). 1 argument used.');
		// Case #13: Reference link. Reference to a cell with string value. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to a cell with string value. Returns number of sheets (1). 1 argument used.');
		// Case #14: Area. Multi-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(A1:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A1:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area. Multi-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #15: Table. Table structured reference with text values. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table structured reference with text values. Returns number of sheets (1). 1 argument used.');
		// Case #16: Name. Named range referring to a range. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name. Named range referring to a range. Returns number of sheets (1). 1 argument used.');
		// Case #17: Name3D. 3D named range referring to a range. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name3D. 3D named range referring to a range. Returns number of sheets (1). 1 argument used.');
		// Case #18: Ref3D. 3D reference to a single cell. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(Sheet1:Sheet5!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Sheet1:Sheet5!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Ref3D. 3D reference to a single cell. Returns number of sheets (1). 1 argument used.');
		// Case #19: Formula. SHEETS inside SUM formula, returns sum of sheet numbers (1). 1 argument used.
		oParser = new parserFormula('SUM(SHEETS(A1:A3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(SHEETS(A1:A3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. SHEETS inside SUM formula, returns sum of sheet numbers (1). 1 argument used.');
		// Case #20: Area3D. 3D multi-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(Sheet2!A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Sheet2!A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area3D. 3D multi-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #21: Formula. Formula resolving to a valid cell reference. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(ADDRESS(1,1,4,,"Sheet2"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(ADDRESS(1,1,4,,"Sheet2")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Formula. Formula resolving to a valid cell reference. Returns number of sheets (1). 1 argument used.');
		// Case #22: Reference link. Reference to a cell in the first sheet. Returns number of sheets (1). 1 argument used.
		oParser = new parserFormula('SHEETS(A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to a cell in the first sheet. Returns number of sheets (1). 1 argument used.');

		// Negative cases:
		// Case #1: String. String not a valid reference (sheet name not accepted). Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS("Sheet2")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS("Sheet2") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String. String not a valid reference (sheet name not accepted). Returns #REF!. 1 argument used.');
		// Case #2: Number. Numeric value not a valid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number. Numeric value not a valid reference. Returns #REF!. 1 argument used.');
		// Case #3: Error. Error input propagates #N/A. 1 argument used.
		oParser = new parserFormula('SHEETS(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error input propagates #N/A. 1 argument used.');
		// Case #4: Empty. Reference to empty cell. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Empty. Reference to empty cell. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #5: Name. Named range with text value. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name. Named range with text value. Returns #REF!. 1 argument used.');
		// Case #6: Name3D. 3D named range with text value. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(TestName3DArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TestName3DArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Name3D. 3D named range with text value. Returns #REF!. 1 argument used.');
		// Case #7: Ref3D. 3D reference to non-existent cell. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(Sheet2!Z1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Sheet2!Z1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Ref3D. 3D reference to non-existent cell. Returns #REF!. 1 argument used.');
		// Case #8: Array. Array constant not a valid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS({1;2;3})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS({1;2;3}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Array. Array constant not a valid reference. Returns #REF!. 1 argument used.');
		// Case #9: Formula. Formula resulting in #NUM! propagates error. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. Returns #REF!. 1 argument used.');
		// Case #10: Boolean. Boolean value not a valid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Boolean. Boolean value not a valid reference. Returns #REF!. 1 argument used.');
		// Case #11: Reference link. Reference to cell with error value (#DIV/0!). Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Reference link. Reference to cell with error value (#DIV/0!). Returns #REF!. 1 argument used.');
		// Case #12: String. Date string not a valid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS("01.01.2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS("01.01.2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String. Date string not a valid reference. Returns #REF!. 1 argument used.');
		// Case #13: Formula. Formula resolving to invalid date (#NUM!). Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(DATE(2025,13,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(DATE(2025,13,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula. Formula resolving to invalid date (#NUM!). Returns #REF!. 1 argument used.');
		// Case #14: Name. Named range referring to invalid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name. Named range referring to invalid reference. Returns #REF!. 1 argument used.');
		// Case #15: Name3D. 3D named range referring to invalid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name3D. 3D named range referring to invalid reference. Returns #REF!. 1 argument used.');
		// Case #16: Table. Table column with invalid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Table. Table column with invalid reference. Returns #REF!. 1 argument used.');
		// Case #17: Formula. Formula resolving to invalid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(INDIRECT("Z1"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(INDIRECT("Z1")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Formula. Formula resolving to invalid reference. Returns #REF!. 1 argument used.');
		// Case #18: Ref3D. 3D reference to invalid cell. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Ref3D. 3D reference to invalid cell. Returns #REF!. 1 argument used.');
		// Case #19: Area. Multi-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.
		oParser = new parserFormula('SHEETS(A101:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A101:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Area. Multi-cell range. Returns number of sheets (1) for the referenced sheet. 1 argument used.');
		// Case #20: Date. Date value not a valid reference. Returns #REF!. 1 argument used.
		oParser = new parserFormula('SHEETS(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Date. Date value not a valid reference. Returns #REF!. 1 argument used.');

		// Bounded cases:
		// Case #1: Reference link. Reference to cell in first sheet (Sheet1). Returns 1. 1 argument used.
		oParser = new parserFormula('SHEETS(A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Reference link. Reference to cell in first sheet (Sheet1). Returns 1. 1 argument used.');
		// Case #2: Ref3D. Reference to cell in second sheet (Sheet2/3D). Returns 1. 1 argument used.
		oParser = new parserFormula('SHEETS(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Ref3D. Reference to cell in second sheet (Sheet2/3D). Returns 1. 1 argument used.');
		// Case #3: Area. Single-cell range in first sheet. Returns 1. 1 argument used.
		oParser = new parserFormula('SHEETS(A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Area. Single-cell range in first sheet. Returns 1. 1 argument used.');
		// Case #4: Area3D. Single-cell 3D range in second sheet. Returns 1. 1 argument used.
		oParser = new parserFormula('SHEETS(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SHEETS(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Area3D. Single-cell 3D range in second sheet. Returns 1. 1 argument used.');

		// Delete created worksheets for tests
		wb.removeWorksheet(0);
		wb.removeWorksheet(0);
		wb.removeWorksheet(0);

		testArrayFormula2(assert, "SHEETS", 1, 1, null, true);
	});

	QUnit.test("Test: \"TYPE\"", function (assert) {
		ws.getRange2("A2").setValue("Smith");

		oParser = new parserFormula("TYPE(A2)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula('TYPE("Mr. "&A2)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula('TYPE(2+A2)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 16);

		oParser = new parserFormula('(2+A2)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('TYPE({1,2;3,4})', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 64);

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
		// Case #1: Number. Basic number input. Returns 1.
		oParser = new parserFormula('TYPE(42)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(42) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number. Basic number input. Returns 1.');
		// Case #2: String. Basic text input. Returns 2.
		oParser = new parserFormula('TYPE("Hello")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE("Hello") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String. Basic text input. Returns 2.');
		// Case #4: Error. Error value #N/A. Returns 16.
		oParser = new parserFormula('TYPE(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Positive case: Error. Error value #N/A. Returns 16.');
		// Case #5: Array. Single-row array. Returns 64.
		oParser = new parserFormula('TYPE({1,2,3})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE({1,2,3}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 64, 'Test: Positive case: Array. Single-row array. Returns 64.');
		// Case #6: Formula. Nested formula returning number. Returns 1.
		oParser = new parserFormula('TYPE(SUM(1,2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(SUM(1,2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Nested formula returning number. Returns 1.');
		// Case #7: Formula. Nested IF returning text. Returns 2.
		oParser = new parserFormula('TYPE(IF(TRUE,"text",1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(IF(TRUE,"text",1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. Nested IF returning text. Returns 2.');
		// Case #8: Date. Date as serial number. Returns 1.
		oParser = new parserFormula('TYPE(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Date. Date as serial number. Returns 1.');
		// Case #9: Time. Time as decimal number. Returns 1.
		oParser = new parserFormula('TYPE(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Time. Time as decimal number. Returns 1.');
		// Case #10: Reference link. Reference to cell with number. Returns 1.
		oParser = new parserFormula('TYPE(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to cell with number. Returns 1.');
		// Case #11: Reference link. Reference to cell with text. Returns 2.
		oParser = new parserFormula('TYPE(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to cell with text. Returns 2.');
		// Case #12: Area. Single-cell range with number. Returns 1.
		oParser = new parserFormula('TYPE(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area. Single-cell range with number. Returns 1.');
		// Case #13: Array. Array with logical values. Returns 64.
		oParser = new parserFormula('TYPE({TRUE,FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE({TRUE,FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 64, 'Test: Positive case: Array. Array with logical values. Returns 64.');
		// Case #14: Name. Named range with number. Returns 1.
		oParser = new parserFormula('TYPE(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name. Named range with number. Returns 1.');
		// Case #15: Name3D. 3D named range with text. Returns 2.
		oParser = new parserFormula('TYPE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name3D. 3D named range with text. Returns 2.');
		// Case #16: Ref3D. 3D reference to cell with number. Returns 1.
		oParser = new parserFormula('TYPE(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. 3D reference to cell with number. Returns 1.');
		// Case #17: Area3D. 3D single-cell range with text. Returns 2.
		oParser = new parserFormula('TYPE(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area3D. 3D single-cell range with text. Returns 2.');
		// Case #18: Table. Table structured reference with number. Returns 1.
		oParser = new parserFormula('TYPE(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table structured reference with number. Returns 1.');
		// Case #19: Formula. Nested formula returning text. Returns 2.
		oParser = new parserFormula('TYPE(CONCAT("A","B"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(CONCAT("A","B")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. Nested formula returning text. Returns 2.');
		// Case #20: Formula. Nested formula with array returning number. Returns 1.
		oParser = new parserFormula('TYPE(AVERAGE({1,2,3}))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(AVERAGE({1,2,3})) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Nested formula with array returning number. Returns 1.');
		// Case #21: Formula. TYPE as part of SUM formula, number input. Returns 2 (1+1).
		oParser = new parserFormula('SUM(TYPE(42),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(TYPE(42),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. TYPE as part of SUM formula, number input. Returns 2 (1+1).');

		// Negative cases:
		// Case #1: Area. Multi-cell range. Returns #VALUE!.
		oParser = new parserFormula('TYPE(A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(A100:A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 64, 'Test: Negative case: Area. Multi-cell range. Returns #VALUE!.');
		// Case #2: Area3D. 3D multi-cell range. Returns #VALUE!.
		oParser = new parserFormula('TYPE(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(Sheet2!A1:A2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 64, 'Test: Negative case: Area3D. 3D multi-cell range. Returns #VALUE!.');
		// Case #3: Empty. Reference to empty cell. Returns #VALUE!.
		oParser = new parserFormula('TYPE(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Empty. Reference to empty cell. Returns #VALUE!.');
		// Case #4: String. Empty string. Returns 2.
		oParser = new parserFormula('TYPE("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: String. Empty string. Returns 2.');
		// Case #5: Name. Named range with multiple cells. Returns #VALUE!.
		oParser = new parserFormula('TYPE(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(TestNameArea2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 64, 'Test: Negative case: Name. Named range with multiple cells. Returns #VALUE!.');
		// Case #6: Name3D. 3D named range with multiple cells. Returns #VALUE!.
		oParser = new parserFormula('TYPE(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(TestNameArea3D2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Negative case: Name3D. 3D named range with multiple cells. Returns #VALUE!.');
		// Case #7: Table. Table column with multiple values. Returns #VALUE!.
		oParser = new parserFormula('TYPE(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: Table. Table column with multiple values. Returns #VALUE!.');
		// Case #8: Formula. Nested formula returning #NUM!. Returns 16.
		oParser = new parserFormula('TYPE(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Negative case: Formula. Nested formula returning #NUM!. Returns 16.');
		// Case #9: Ref3D. 3D reference to cell with error. Returns 16.
		oParser = new parserFormula('TYPE(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: Ref3D. 3D reference to cell with error. Returns 16.');
		// Case #10: Reference link. Reference to cell with error. Returns 16.
		oParser = new parserFormula('TYPE(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Reference link. Reference to cell with error. Returns 16.');
		// Case #11: Number. Number too small, causes #NUM!. Returns 16.
		oParser = new parserFormula('TYPE(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Number. Number too small, causes #NUM!. Returns 16.');
		// Case #12: Formula. Formula resulting in #DIV/0!. Returns 16.
		oParser = new parserFormula('TYPE(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Negative case: Formula. Formula resulting in #DIV/0!. Returns 16.');
		// Case #13: Array. Mixed array (number and text). Returns 64.
		oParser = new parserFormula('TYPE({1,"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE({1,"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 64, 'Test: Negative case: Array. Mixed array (number and text). Returns 64.');
		// Case #14: Formula. Nested formula returning #VALUE!. Returns 16.
		oParser = new parserFormula('TYPE(CHOOSE(0,"A","B"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(CHOOSE(0,"A","B")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Negative case: Formula. Nested formula returning #VALUE!. Returns 16.');
		// Case #15: Reference link. Reference to cell with boolean. Returns 4.
		oParser = new parserFormula('TYPE(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Reference link. Reference to cell with boolean. Returns 4.');
		// Case #16: Time. Invalid time value. Returns 16.
		oParser = new parserFormula('TYPE(TIME(25,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(TIME(25,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Time. Invalid time value. Returns 16.');
		// Case #17: Date. Date before Excelâ??s base date. Returns 16.
		oParser = new parserFormula('TYPE(DATE(1899,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(DATE(1899,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Date. Date before Excelâ??s base date. Returns 16.');
		// Case #18: String. String not convertible to number. Returns 2.
		oParser = new parserFormula('TYPE("1/0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE("1/0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: String. String not convertible to number. Returns 2.');
		// Case #19: Formula. Formula resulting in overflow. Returns 16.
		oParser = new parserFormula('TYPE(1E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(1E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Formula. Formula resulting in overflow. Returns 16.');
		// Case #20: Array. Array with error value. Returns 64.
		oParser = new parserFormula('TYPE({#N/A})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE({#N/A}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 64, 'Test: Negative case: Array. Array with error value. Returns 64.');

		// Bounded cases:
		// Case #1: Number. Smallest positive number. Returns 1.
		oParser = new parserFormula('TYPE(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Smallest positive number. Returns 1.');
		// Case #2: Number. Largest valid Excel number. Returns 1.
		oParser = new parserFormula('TYPE(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Largest valid Excel number. Returns 1.');
		// Case #3: Date. Earliest valid Excel date. Returns 1.
		oParser = new parserFormula('TYPE(DATE(1900,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(DATE(1900,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Date. Earliest valid Excel date. Returns 1.');
		// Case #4: Date. Latest valid Excel date. Returns 1.
		oParser = new parserFormula('TYPE(DATE(9999,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TYPE(DATE(9999,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Date. Latest valid Excel date. Returns 1.');

		// Need to fix: area,ref,name handle
		// Case #1: Area. Multi-cell range. Returns #VALUE!.
		// Case #2: Area3D. 3D multi-cell range. Returns #VALUE!.
		// Case #3: Empty. Reference to empty cell. Returns #VALUE!.
		// Case #5: Name. Named range with multiple cells. Returns #VALUE!.
		// Case #6: Name3D. 3D named range with multiple cells. Returns #VALUE!.


		//TODO нужна другая функция для тестирования
		//testArrayFormula2(assert, "TYPE", 1, 1);
	});

	wb.dependencyFormulas.unlockRecal();
});
