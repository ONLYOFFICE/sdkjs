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
$(function () {
	let documents = {
		"test_formats": Asc.test_formats,
		"test_formats_redo": Asc.test_formats_redo,
		"test_formats2": Asc.test_formats,
		"test_formats2_redo": Asc.test_formats_redo,
	}
	let files = {};
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
	api.initCollaborativeEditing({});
	window["Asc"]["editor"] = api;

	waitLoadModules(function () {
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		startTests();
	});

	function waitLoadModules(waitCallback) {
		Asc.spreadsheet_api.prototype._init = function () {
			this._loadModules();
		};
		Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
			callback();
		};
		Asc.spreadsheet_api.prototype.onEndLoadFile = function (fonts, callback) {
			waitCallback();
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
		AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
		};
		AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
		};
		AscCommonExcel.WorksheetView.prototype.setSelection = function () {
		};
		AscCommonExcel.WorksheetView.prototype.draw = function () {
		};
		AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
		};
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
		};
	}

	function openDocument(file) {
		if (api.wbModel) {
			api.asc_CloseFile();
		}

		api.isOpenOOXInBrowser = false;
		api.openingEnd.xlsx = true;
		api.openingEnd.data = AscCommon.Base64.decode(file["Editor.xlsx"]);
		api._openDocument(AscCommon.Base64.decode(file["Editor.bin"]));
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		return api.wbModel;
	}

	function prepareTest(assert, wb) {
		api.wb.model = wb;
		api.wbModel = wb;
		api.initGlobalObjects(wb);
		api.handlers.remove("getSelectionState");
		api.handlers.add("getSelectionState", function () {
			return null;
		});
		api.handlers.remove("asc_onError");
		api.handlers.add("asc_onError", function (code, level) {
			assert.equal(code, 0, "asc_onError");
		});
		AscCommon.History.Clear();
	}

	let memory = new AscCommon.CMemory();

	function Utf8ArrayToStr(array) {
		let out, i, len, c;
		let char2, char3;

		out = "";
		len = array.length;
		i = 0;
		while (i < len) {
			c = array[i++];
			switch (c >> 4) {
				case 0:
				case 1:
				case 2:
				case 3:
				case 4:
				case 5:
				case 6:
				case 7:
					// 0xxxxxxx
					out += String.fromCharCode(c);
					break;
				case 12:
				case 13:
					// 110x xxxx   10xx xxxx
					char2 = array[i++];
					out += String.fromCharCode(((c & 0x1F) << 6) | (char2 & 0x3F));
					break;
				case 14:
					// 1110 xxxx  10xx xxxx  10xx xxxx
					char2 = array[i++];
					char3 = array[i++];
					out += String.fromCharCode(((c & 0x0F) << 12) |
						((char2 & 0x3F) << 6) |
						((char3 & 0x3F) << 0));
					break;
			}
		}

		return out;
	}

	function getXml(pivot, addCacheDefinition) {
		memory.Seek(0);
		pivot.toXml(memory);
		if (addCacheDefinition) {
			memory.WriteXmlString('\n\n');
			pivot.cacheDefinition.toXml(memory);
		}
		let buffer = new Uint8Array(memory.GetCurPosition());
		for (let i = 0; i < memory.GetCurPosition(); i++) {
			buffer[i] = memory.data[i];
		}
		if (typeof TextDecoder !== "undefined") {
			return new TextDecoder("utf-8").decode(buffer);
		} else {
			return Utf8ArrayToStr(buffer);
		}

	}

	function checkHistoryOperation(assert, pivot, valuesUndo, valuesRedo, message, action, check) {
		let ws = pivot.GetWS();
		let wb = ws.workbook;
		let xmlUndo = getXml(pivot, false);
		let pivotStart = pivot.clone();
		pivotStart.Id = pivot.Get_Id();

		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		action();
		AscCommon.History.EndTransaction();
		pivot = wb.getPivotTableById(pivot.Get_Id());
		check(assert, pivot, valuesRedo, message);
		let xmlDo = getXml(pivot, true);
		let changes = wb.SerializeHistory()[0];

		AscCommon.History.Undo();
		pivot = wb.getPivotTableById(pivot.Get_Id());
		check(assert, pivot, valuesUndo, message + "_undo");
		assert.strictEqual(getXml(pivot, false), xmlUndo, message + "_undo_xml");

		AscCommon.History.Redo();
		pivot = wb.getPivotTableById(pivot.Get_Id());
		check(assert, pivot, valuesRedo, message + "_redo");
		assert.strictEqual(getXml(pivot, true), xmlDo, message + "_redo_xml");

		AscCommon.History.Undo();
		ws.deletePivotTable(pivot.Get_Id());
		pivot = pivotStart;
		ws.insertPivotTable(pivot, false, false);
		wb.DeserializeHistory(changes);
		pivot = wb.getPivotTableById(pivot.Get_Id());
		check(assert, pivot, valuesRedo, message + "_changes");
		assert.strictEqual(getXml(pivot, true), xmlDo, message + "_changes_xml");
		return pivot;
	}

	function getReportValues(pivot) {
		let res = [];
		let range = new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange();
		pivot.GetWS().getRange3(range.r1, range.c1, range.r2, range.c2)._foreach(function (cell, r, c, r1, c1) {
			if (!res[r - r1]) {
				res[r - r1] = [];
			}
			res[r - r1][c - c1] = cell.getName() + ":" + cell.getValue();
		});
		return res;
	}

	function getReportValuesWithBoolFill(pivot) {
		let res = [];
		let range = new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange();
		pivot.GetWS().getRange3(range.r1, range.c1, range.r2, range.c2)._foreach(function (cell, r, c, r1, c1) {
			if (!res[r - r1]) {
				res[r - r1] = [];
			}
			res[r - r1][c - c1] = cell.getName() + ":" + cell.getValue() + ":" + !!(cell.getStyle() && !cell.getStyle().isNormalFill());
		});
		return res;
	}

	function getReportValuesWithBoolFillAndNum(pivot) {
		let res = [];
		let range = new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange();
		pivot.GetWS().getRange3(range.r1, range.c1, range.r2, range.c2)._foreach(function (cell, r, c, r1, c1) {
			if (!res[r - r1]) {
				res[r - r1] = [];
			}
			let xf = cell.getStyle();
			res[r - r1][c - c1] = cell.getName() + ":" + cell.getValue() + ":" + !!(xf && (!xf.isNormalFill() || !xf.num || !xf.num.getNumFormat().isGeneralFormat()));
		});
		return res;
	}

	QUnit.module("Pivot");

	function startTests() {
		QUnit.start();
		QUnit.test('Test: refresh test_formats check values and format', function (assert) {
			let file = Asc.test_formats;
			let fileRedo = Asc.test_formats_redo;
			let wsNames = [
				"Default data",
				"Default Label",
				"MultidataField default data",
				"MultidataField default label",
				"Multidatafield label offset",
				"TopRight&Origin",
				"Buttons compact test",
				"Buttons tabular test",
				"Buttons outline test",
				"Filter test",
				"TypeAll test",
				"NoDataFieldAllTest",
				"All test"
			];
			let row = 4;
			let col = 0;
			let getValues = getReportValuesWithBoolFill;

			function prepareValues(wb, name, row, col) {
				let pivot = wb.getWorksheetByName(name).getPivotTable(col, row);
				return getValues(pivot);
			}

			function preparePivots(wb, name, row, col) {
				return wb.getWorksheetByName(name).getPivotTable(col, row);
			}

			let wbRedo = openDocument(fileRedo);
			let valuesRedo = wsNames.map(function (name) {
				return prepareValues(wbRedo, name, row, col);
			});

			let wb = openDocument(file);
			let valuesUndo = wsNames.map(function (name) {
				return prepareValues(wb, name, row, col);
			});
			let pivots = wsNames.map(function (name) {
				return preparePivots(wb, name, row, col);
			});

			prepareTest(assert, wb);
			wsNames.forEach(function (name, index) {
				let pivot = pivots[index];
				pivot = checkHistoryOperation(assert, pivot, valuesUndo[index], valuesRedo[index], "refresh[" + name + "]", function () {
					pivot.asc_refresh(api);
				}, function (assert, pivot, values, message) {
					assert.deepEqual(getValues(pivot), values, message);
				});
			});
		});
		QUnit.test('Test: refresh pivot-styles-numformat check values and format', function (assert) {
			let file = Asc.pivot_styles_numformat;
			let fileRedo = Asc.pivot_styles_numformat_redo;
			let wsNames = [
				"pivot"
			];
			let row = 4;
			let col = 0;
			let getValues = getReportValuesWithBoolFillAndNum;

			function prepareValues(wb, name, row, col) {
				let pivot = wb.getWorksheetByName(name).getPivotTable(col, row);
				return getValues(pivot);
			}

			function preparePivots(wb, name, row, col) {
				return wb.getWorksheetByName(name).getPivotTable(col, row);
			}

			let wbRedo = openDocument(fileRedo);
			let valuesRedo = wsNames.map(function (name) {
				return prepareValues(wbRedo, name, row, col);
			});

			let wb = openDocument(file);
			let valuesUndo = wsNames.map(function (name) {
				return prepareValues(wb, name, row, col);
			});
			let pivots = wsNames.map(function (name) {
				return preparePivots(wb, name, row, col);
			});

			prepareTest(assert, wb);
			wsNames.forEach(function (name, index) {
				let pivot = pivots[index];
				pivot = checkHistoryOperation(assert, pivot, valuesUndo[index], valuesRedo[index], "refresh[" + name + "]", function () {
					pivot.asc_refresh(api);
				}, function (assert, pivot, values, message) {
					assert.deepEqual(getValues(pivot), values, message);
				});
			});
		});
		QUnit.test('Test: refresh pivot check formats after dataField reindex', function (assert) {
			const file = Asc.pivot_datafield_reindex;
			const row = 4;
			const col = 0;
			const wb = openDocument(file);
			const getValues = getReportValuesWithBoolFillAndNum;

			const dataFieldStartPivot = wb.getWorksheetByName('DataFieldStart').getPivotTable(col, row);
			const moveDataFieldResultPivot = wb.getWorksheetByName('moveDataFieldResult').getPivotTable(col, row);
			const removeDataFieldResultPivot = wb.getWorksheetByName('RemoveDataFieldResult').getPivotTable(col, row);

			const dataFieldStartValues = getValues(dataFieldStartPivot);
			const moveDataFieldResultValues = getValues(moveDataFieldResultPivot);
			const removeDataFieldResultValues = getValues(removeDataFieldResultPivot);

			prepareTest(assert, wb);
			let pivot = wb.getWorksheetByName('DataFieldStart').getPivotTable(col, row);
			pivot = checkHistoryOperation(assert, dataFieldStartPivot, dataFieldStartValues, moveDataFieldResultValues, "move dataField", function () {
				pivot.asc_moveDataField(api, 0, 1);
				pivot.asc_refresh(api);
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getValues(pivot), values, message);
			});
			pivot = checkHistoryOperation(assert, dataFieldStartPivot, moveDataFieldResultValues, removeDataFieldResultValues, "remove dataField", function () {
				pivot.asc_removeDataField(api, 5, 1);
				pivot.asc_refresh(api);
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getValues(pivot), values, message);
			});
		});
		QUnit.test('Test: refresh pivot check formats after change source data 1', function (assert) {
			const getValues = getReportValuesWithBoolFillAndNum;
			const file = Asc.pivot_data_reindex1;
			const row = 4;
			const col = 0;
			const wb = openDocument(file);
			const dataRefFieldSettings = 'Sheet1' + "!" + 'I2:O13';
			let pivot = wb.getWorksheetByName('pivot').getPivotTable(col, row);
			const values = getValues(pivot);
			prepareTest(assert, wb);

			pivot = checkHistoryOperation(assert, pivot, values, values, "change source data 1", function () {
				const props = new Asc.CT_pivotTableDefinition();
				props.asc_setDataRef(dataRefFieldSettings);
				pivot.asc_set(api, props);
				pivot.asc_refresh(api);
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getValues(pivot), values, message);
			});
		});
		QUnit.test('Test: refresh pivot check formats after change source data 2', function (assert) {
			const file = Asc.pivot_data_reindex2;
			const row = 4;
			const col = 0;
			const wb = openDocument(file);
			const dataRefFieldSettings1 = 'Sheet1' + "!" + 'I2:O8';
			const dataRefFieldSettings2 = 'Sheet1' + "!" + 'A2:G13';
			const getValues = getReportValuesWithBoolFillAndNum;

			function prepareValues(wb, name, row, col) {
				let pivot = wb.getWorksheetByName(name).getPivotTable(col, row);
				return getValues(pivot);
			}

			const valuesStart = prepareValues(wb, 'start', row, col);
			const valuesReindex = prepareValues(wb, 'reindex', row, col);
			const valuesResult = prepareValues(wb, 'result', row, col);
			let pivot = wb.getWorksheetByName('start').getPivotTable(col, row);
			prepareTest(assert, wb);

			pivot = checkHistoryOperation(assert, pivot, valuesStart, valuesReindex, "change source data 2 reindex", function () {
				const props = new Asc.CT_pivotTableDefinition();
				props.asc_setDataRef(dataRefFieldSettings1);
				pivot.asc_set(api, props);
				pivot.asc_refresh(api);
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getValues(pivot), values, message);
			});
			pivot = checkHistoryOperation(assert, pivot, valuesReindex, valuesResult, "change source data 2 result", function () {
				const props = new Asc.CT_pivotTableDefinition();
				props.asc_setDataRef(dataRefFieldSettings2);
				pivot.asc_set(api, props);
				pivot.asc_refresh(api);
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getValues(pivot), values, message);
			});
		});
		QUnit.test('Test: refresh pivot check formats after remove field', function (assert) {
			const file = Asc.remove_field;
			const row = 4;
			const col = 0;
			const wb = openDocument(file);
			const getValues = getReportValuesWithBoolFillAndNum;

			function prepareValues(wb, name, row, col) {
				let pivot = wb.getWorksheetByName(name).getPivotTable(col, row);
				return getValues(pivot);
			}

			const valuesStart = prepareValues(wb, 'start', row, col);
			const valuesResult = prepareValues(wb, 'result', row, col);
			let pivot = wb.getWorksheetByName('start').getPivotTable(col, row);
			prepareTest(assert, wb);

			pivot = checkHistoryOperation(assert, pivot, valuesStart, valuesResult, "remove field", function () {
				pivot.asc_removeField(api, 0);
				pivot.asc_addRowField(api, 0);
				pivot.asc_moveRowField(api, 1, 0);
				pivot.asc_refresh(api);
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getValues(pivot), values, message);
			});
		});
		QUnit.test('Test: GETPIVOTDATA', function (assert) {
			const file = Asc.GetPivotData;
			const wb = openDocument(file);
			let data = {
				"general": [
					['G4', 'E11'],
					['G5', 'E8'],
					['G6', 'E9'],
					['G7', 'D9'],
					['G8', 'D11'],
					['G9', 'B10'],
					['G10', 'E11', true],
					['H4', 'E11', true],
					['H5', 'D9', true],
					['H6', 'E11', true],
					['H7', 'D9', true],
					['I4', 'E11', true],
					['I5', 'E8', true],
					['I6', 'E8', true],
					['I7', 'E11', true],
					['G15', 'C17'],
					['G16', 'B17'],
					['G17', 'C17', true],
					['G21', 'A20', true],
					['H21', 'A20', true],
					['I21', 'A20', true],
					['I22', 'A20', true],
					['I23', 'B24'],
					['G28', 'B34'],
					['G29', 'B41'],
					['G30', 'B41', true],
					['G46', 'A45', true],
					['H46', 'A45', true],
				],
				"total": [
					['H3', 'D7'],
					['H10', 'B13'],
					['H17', 'D19'],
					['H24', 'A24', true],
					['H31', 'A31', true],
					['H37', 'A37', true],
					['H43', 'A44'],
				],
				"subtotal": [
					['D3', 'B3'],
					['D10', 'B10'],
					['D18', 'B18', true],
					['D21', 'B21'],
					['D25', 'B25', true],
					['D26', 'B26'],
					['D34', 'B34'],
					['K3', 'H3'],
					['K10', 'H10'],
					['K18', 'H18', true],
					['K21', 'H21'],
					['K25', 'H25', true],
					['K26', 'I26'],
					['K34', 'I34'],
					['R5', 'O5'],
					['R12', 'O12'],
					['R18', 'O18', true],
					['R25', 'O25', true],
					['R28', 'P28'],
					['R33', 'P33'],
					['AB5', 'Y5'],
					['AB11', 'V9', true],
					['AB12', 'Y12'],
					['AH17', 'V16', true],
					['AH18', 'Z20'],
					['AH20', 'AE20'],
				],
				"values-col": [
					['AG7', 'Y7'],
					['AH7', 'AC7'],
					['AI7', 'AE7'],
					['AG13', 'Y13'],
					['AH13', 'AC13'],
					['AI13', 'AE13'],
					['AG20', 'X20'],
					['AH20', 'AC20'],
					['AI20', 'AE20'],
					['AG26', 'X26'],
					['AH26', 'AC26'],
					['AI26', 'AE26'],
					['AG33', 'Y33'],
					['AH33', 'AC33'],
					['AI33', 'AE33'],
					['AG39', 'Y39'],
					['AH39', 'AC39'],
					['AI39', 'AE39'],
				],
				"values-row": [
					['R7', 'K7'],
					['S7', 'O7'],
					['T7', 'P7'],
					['R20', 'K20'],
					['S20', 'O20'],
					['T20', 'P20'],
					['R30', 'K30'],
					['S30', 'O30'],
					['T30', 'P30'],
					['R45', 'K45'],
					['S45', 'O45'],
					['T45', 'P45'],
					['R57', 'K57'],
					['S57', 'O57'],
					['T57', 'P57'],
					['R70', 'K70'],
					['S70', 'O70'],
					['T70', 'P70'],
				],
				"groups": [
					['F5', 'D5'],
					['F6', 'D6'],
					['F7', 'D7'],
					['F8', 'D8'],
					['F9', 'D9'],
					['F10', 'D10'],
					['F11', 'D11'],
					['G5', 'C6'],
					['G8', 'C8'],
					['F15', 'D15'],
					['F16', 'D16'],
					['F17', 'D17'],
					['F18', 'D18'],
					['F19', 'D19'],
					['F20', 'D20'],
					['F21', 'D21'],
					['F22', 'D22'],
					['F23', 'D23'],
					['F24', 'D24'],
					['F15', 'D15'],
					['G15', 'B15'],
					['G16', 'B16'],
					['G17', 'B17'],
					['G18', 'B18'],
					['G19', 'B19'],
					['G20', 'B20'],
					['F35', 'D35'],
				]
			};
			for (let sheetName in data) {
				let elems = data[sheetName];
				let ws = wb.getWorksheetByName(sheetName);
				elems.forEach(function (elem) {
					let errorText = sheetName + ':' + JSON.stringify(elem);
					let formulaRef = elem[0];
					let rangeFormula = ws.getRange2(formulaRef);
					let valueExpected = rangeFormula.getValue();
					let formulaExpected = rangeFormula.getValueForEdit().substring(1);

					if (!elem[2]) {
						let pivotRef = elem[1];
						let rangePivot = ws.getRange2(pivotRef);
						let bboxPivot = rangePivot.bbox;
						let pivot = ws.getPivotTable(bboxPivot.c1, bboxPivot.r1);
						let formula = pivot.getGetPivotDataFormulaByActiveCell(bboxPivot.r1, bboxPivot.c1, false);
						assert.strictEqual(formula, formulaExpected, errorText);
					}

					let oParser = new AscCommonExcel.parserFormula(formulaExpected, "A1", ws);
					assert.ok(oParser.parse(), errorText);
					assert.strictEqual(oParser.calculate().getValue() + "", valueExpected, errorText);
				});
			}

		});
		QUnit.test('Test: GETPIVOTDATA TWO ARGS', function (assert) {
			const file = Asc.GetPivotData2;
			const wb = openDocument(file);
			const data = {
				'General': [
					'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9', 'G10', 'G11', 'G12', 'G13', 'G14', 'G15', 'G16', 'G17', 'G18', 'G19',
					'G22', 'G23', 'G24', 'G25', 'G30',
					'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'H9', 'H10', 'H11',
					'I3', 'I22', 'I23', 'I24', 'I25'
				],
				'Subtotals + GrandTotals': [
					'G5', 'G6', 'G14', 'G15', 'G38', 'G39', 'G49', 'G50', 'G51', 'G55', 'G56', 'G57',
					'H24', 'H25', 'H38', 'H39', 'H49', 'H50', 'H55', 'H56'
				],
				'DataFields': [
					'I4', 'I5', 'I6', 'I7', 'I8', 'I9',
					'J4', 'J5', 'J6', 'J7', 'J8', 'J9', 'J10'
				]
			}
			for (let sheetName in data) {
				const elems = data[sheetName];
				const ws = wb.getWorksheetByName(sheetName);
				elems.forEach(function (formulaRef) {
					let errorText = sheetName + ':' + formulaRef;
					let rangeFormula = ws.getRange2(formulaRef);
					let valueExpected = rangeFormula.getValue();
					let formula = rangeFormula.getValueForEdit().substring(1);

					let oParser = new AscCommonExcel.parserFormula(formula, "A1", ws);
					assert.ok(oParser.parse(), errorText);
					const value = oParser.calculate().getValue();
					assert.strictEqual(value + "", valueExpected, errorText);
				});
			}
		});
		QUnit.test('Test: Api pivot builder', function (assert) {
			const file = Asc.pivotBuilder;
			const wb = openDocument(file);
			ws = wb.getWorksheetByName('Test');
			const pivot = wb.getWorksheetByName('Test').getPivotTable(2, 3);

			let range = pivot.asc_getColumnRange();
			assert.deepEqual([range.bbox.r1, range.bbox.r2, range.bbox.c1, range.bbox.c2], [2, 3, 1, 3], 'column range');

			range = pivot.asc_getRowRange();
			assert.deepEqual([range.bbox.r1, range.bbox.r2, range.bbox.c1, range.bbox.c2], [3, 12, 0, 0], 'row range');

			range = pivot.asc_getDataBodyRange();
			assert.deepEqual([range.bbox.r1, range.bbox.r2, range.bbox.c1, range.bbox.c2], [4, 12, 1, 3], 'data range');

			let getDataToGetPivotData = pivot.asc_getDataToGetPivotData(['East', 'Boy', 'Fancy']);
			let params = pivot.getGetPivotParamsByActiveCell({row: 5, col: 1})
			assert.deepEqual(getDataToGetPivotData, {
				dataFieldName: 'Sum of Price',
				optParams: params.optParams
			}, 'getGetPivotParamsByActiveCell');

		});
		QUnit.test('Test: CALCULATED-ITEMS refresh', function (assert) {
			const file = Asc.test_calculated;
			const wb = openDocument(file);
			const row = 4;
			const col = 0;
			let pivot = wb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			prepareTest(assert, wb);

			const standard = getReportValues(pivot);
			pivot = checkHistoryOperation(assert, pivot, standard, standard, "refresh pivot", function () {
				pivot.asc_refresh(api);
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getReportValues(pivot), values, message);
			});
		});
		QUnit.test('Test: CALCULATED-ITEMS add', function (assert) {
			const fileStart = Asc.AddCalculatedItemsStart;
			const fileEnd = Asc.AddCalculatedItemsStandard;
			const startWb = openDocument(fileStart);
			const standardWb = openDocument(fileEnd);
			const row = 4;
			const col = 0;
			let pivot = startWb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			const standardPivot = standardWb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			prepareTest(assert, startWb);
			const undo = getReportValues(pivot);
			const standard = getReportValues(standardPivot);
			pivot = checkHistoryOperation(assert, pivot, undo, standard, "add items", function () {
				pivot.asc_addCalculatedItem(api, 0, 'Formula3', pivot.asc_convertCalculatedFormula("East2-10+'we''s t'", 0));
				pivot.asc_addCalculatedItem(api, 1, 'Formula2', pivot.asc_convertCalculatedFormula("Boy-Girl", 1));
				pivot.asc_addCalculatedItem(api, 2, 'Formula1', pivot.asc_convertCalculatedFormula("Fancy+Tee", 2));
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getReportValues(pivot), values, message);
			});
		});
		QUnit.test('Test: CALCULATED-ITEMS remove', function (assert) {
			const fileStart = Asc.AddCalculatedItemsStandard;
			const fileEnd = Asc.RemoveCalculatedItemsStandard;
			const startWb = openDocument(fileStart);
			const standardWb = openDocument(fileEnd);
			const row = 4;
			const col = 0;
			let pivot = startWb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			const standardPivot = standardWb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			prepareTest(assert, startWb);
			const undo = getReportValues(pivot);
			const standard = getReportValues(standardPivot);
			pivot = checkHistoryOperation(assert, pivot, undo, standard, "add items", function () {
				pivot.asc_removeCalculatedItem(api, 0, 'Formula1');
				pivot.asc_removeCalculatedItem(api, 1, 'Formula2');
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getReportValues(pivot), values, message);
			});
		});
		QUnit.test('Test: CALCULATED-ITEMS modify', function (assert) {
			const fileStart = Asc.RemoveCalculatedItemsStandard;
			const fileEnd = Asc.ModifyCalculatedItemsStandard;
			const startWb = openDocument(fileStart);
			const standardWb = openDocument(fileEnd);
			const row = 4;
			const col = 0;
			let pivot = startWb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			const standardPivot = standardWb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			prepareTest(assert, startWb);
			const undo = getReportValues(pivot);
			const standard = getReportValues(standardPivot);
			pivot = checkHistoryOperation(assert, pivot, undo, standard, "add items", function () {
				pivot.asc_modifyCalculatedItem(api, 0, 'Formula2', pivot.asc_convertCalculatedFormula("East2-15", 0));
				pivot.asc_modifyCalculatedItem(api, 2, 'Formula1', pivot.asc_convertCalculatedFormula("Fancy+Tee-10", 2));
			}, function (assert, pivot, values, message) {
				assert.deepEqual(getReportValues(pivot), values, message);
			});
		});
		QUnit.test('Test: CALCULATED-ITEMS getFieldIndexByCell', function (assert) {
			const file = Asc.AddCalculatedItemsStandard;
			const wb = openDocument(file);
			const row = 4;
			const col = 0;
			const standardFieldIndexes = [
				[null, 2, null, null, null, null],
				[0, 2, 2, 2, 2, null],
				[0, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[0, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[0, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[0, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[0, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[1, null, null, null, null, null],
				[null, null, null, null, null, null],
			];
			const fieldIndexes = [];
			let pivot = wb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			for (let i = 2; i < 30; i += 1) {
				let row = [];
				for (let j = 0; j < 6; j += 1) {
					row.push(pivot.getFieldIndexByCell(i, j));
				}
				fieldIndexes.push(row);
			}
			assert.deepEqual(fieldIndexes, standardFieldIndexes, 'check getFieldIndexes')
		});
		QUnit.test('Test: CALCULATED-ITEMS hasTablesErrorForCalculatedItems', function (assert) {
			const file = Asc.isTablesValidForCalculatedItems;
			const wb = openDocument(file);
			const pivot1 = wb.getWorksheetByName('Sheet2').getPivotTable(0, 4);
			const pivot2 = wb.getWorksheetByName('Sheet2').getPivotTable(0, 11);
			const pivot3 = wb.getWorksheetByName('Sheet2').getPivotTable(0, 18);
			assert.equal(pivot1.asc_hasTablesErrorForCalculatedItems(0), c_oAscError.ID.NotUniqueFieldWithCalculated);
			assert.equal(pivot2.asc_hasTablesErrorForCalculatedItems(0), c_oAscError.ID.CalculatedItemInPageField);
			assert.equal(pivot3.asc_hasTablesErrorForCalculatedItems(0), c_oAscError.ID.No);
		});
		QUnit.test('Test: CALCULATED-ITEMS can add calculatedItemsName', function (assert) {
			const file = Asc.AddCalculatedItemsStandard;
			const wb = openDocument(file);
			const pivot = wb.getWorksheetByName('Sheet2').getPivotTable(0, 4);
			assert.equal(pivot.asc_canAddNameCalculatedItem(0, 'East2'), false);
			assert.equal(pivot.asc_canAddNameCalculatedItem(0, 'East'), false);
			assert.equal(pivot.asc_canAddNameCalculatedItem(0, 'Formula2'), false);
			assert.equal(pivot.asc_canAddNameCalculatedItem(0, 'East3'), true);
			assert.equal(pivot.asc_canAddNameCalculatedItem(1, 'East3'), true);
			assert.equal(pivot.asc_canAddNameCalculatedItem(1, 'Boy'), false);
		});
		QUnit.test('Test: CALCULATED-ITEMS canChangeCalculatedItemsByActiveCell', function (assert) {
			const file = Asc.AddCalculatedItemsStandard;
			const wb = openDocument(file);
			const row = 4;
			const col = 0;
			const standardCanAdd = [
				[false, true, false, false, false, false],
				[true, true, true, true, true, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[true, false, false, false, false, false],
				[false, false, false, false, false, false],
			];
			const canAdd = [];
			let pivot = wb.getWorksheetByName('Sheet2').getPivotTable(col, row);
			for (let i = 2; i < 30; i += 1) {
				let row = [];
				for (let j = 0; j < 6; j += 1) {
					row.push(pivot.canChangeCalculatedItemByCell(i, j));
				}
				canAdd.push(row);
			}
			assert.deepEqual(canAdd, standardCanAdd, 'check can change calculatedItems')
		});
		QUnit.test('Test: change pivot headers and labels', function (assert) {
			const file = Asc.PivotFieldNames;
			const wb = openDocument(file);
			const data = {
				'General': {
					'Editable': {
						'A1': 'testName_A1',
						'A4': 'testName_A4',
						'C3': 'testName_C3',
						'B4': 'testName_B4',
						'C4': 'testName_C4',
						'D4': 'testName_D4',
						'E4': 'testName_E4',
						'A5': 'testName_A5',
						'B5': 'testName_B5',
						'B6': 'testName_B6',
						'B7': 'testName_B7',
						'B8': 'testName_B8',
						'B9': 'testName_B9',
						'B10': 'testName_B10',
						'B11': 'testName_B11',
						'B12': 'testName_B12',
						'B13': 'testName_B13',
						'A16': 'testName_A16',
						'B16': 'testName_B16',
						'B17': 'testName_B17',
						'B18': 'testName_B18',
						'B19': 'testName_B19',
						'B20': 'testName_B20',
						'B21': 'testName_B21',
						'B22': 'testName_B22',
						'B23': 'testName_B23',
						'B24': 'testName_B24'
					}
				},
				'ValuesRow': {
					'Editable': {
						'A3': 'testName_A3',
						'B3': 'testName_B3',
						'A4': 'testName_A4',
						'B4': 'testName_B4',
						'B5': 'testName_B5',
						'A6': 'testName_A6',
						'B6': 'testName_B6',
						'B7': 'testName_B7',
					}
				},
				'DataOnly': {
					'Editable': {
						'A3': 'testName_A3',
					}
				},
				'Default': {
					'Editable': {
						'A3': 'testName_A3',
						'B3': 'testName_B3',
						'A4': 'testName_A4',
						'B4': 'testName_B4',
						'C4': 'testName_C4',
						'D4': 'testName_D4',
						'A5': 'testName_A5',
						'A6': 'testName_A6',
						'A7': 'testName_A7',

					}
				}
			}
			ws = wb.getWorksheetByName('General');
			let pivot = wb.getWorksheetByName('General').getPivotTable(0, 2);

			const filterRange = Asc.Range(0, 0, 0, 0);
			assert.equal(!!pivot.rangeMapper.getEditCellFunction(filterRange), !!data['General']['Editable']['A1'], 'General:A1');

			for (let sheetName in data) {
				ws = wb.getWorksheetByName(sheetName);
				let pivot = wb.getWorksheetByName(sheetName).getPivotTable(0, 2);
				const editable = data[sheetName]['Editable'];
				const pivotRange = pivot.getRange();

				const range = new AscCommonExcel.Range(ws, pivotRange.r1, pivotRange.c1, pivotRange.r2, pivotRange.c2);
				range._foreach(function (cell, r, c, r1, c1) {
					const eachRange = Asc.Range(c, r, c, r);
					const name = eachRange.getName();
					const func = pivot.rangeMapper.getEditCellFunction(eachRange);
					assert.equal(!!func, !!editable[name], sheetName + ':' + name);
				});
				prepareTest(assert, wb);
				for (let refName in editable) {
					const wsRange = ws.getRange2(refName);
					const bbox = ws.getRange2(refName).bbox;
					const startValue = wsRange.getValue();
					pivot = checkHistoryOperation(assert, pivot, startValue, editable[refName], sheetName + ':' + 'A1' + ':editing', function () {
						const func = pivot.rangeMapper.getEditCellFunction(bbox);
						func(editable[refName]);
						pivot.asc_refresh(api);
					}, function (assert, pivot, value, message) {
						assert.equal(ws.getRange2(refName).getValue(), value, message);
					});
				}
			}
		});
	}
});
