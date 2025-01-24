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

	// Tests completed in 3980 milliseconds.
	// 3322 assertions of 3355 passed, 33 failed.

	// To add new test
	// 1) Create xlsx file in Excel with data and pivot
	// 2) Uncomment code below and put it at the end of PivotTables.js
	// 3) Modify AscCommon.baseEditorsApi.prototype.onDocumentContentReady according to your xlsx
	// 4) Open xlsx in editor and copy from console 'testData' and 'standards' to your new test
	// 5) Use exiting test as template for example 'testPivotMisc'

	// function getValues(ws, range) {
	// 	var res = [];
	// 	ws.getRange3(range.r1, range.c1, range.r2, range.c2)._foreach(function(cell, r, c, r1, c1) {
	// 		if (!res[r - r1]) {
	// 			res[r - r1] = [];
	// 		}
	// 		res[r - r1][c - c1] = cell.getValue();
	// 	});
	// 	return res;
	// }
	// function getReportValues(pivot) {
	// 	pivot.init();
	// 	return getValues(pivot.GetWS(), new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
	// }
	//
	// function getTestMatrix(ws) {
	// 	if (!ws || !ws.pivotTables[0]) {
	// 		return "";
	// 	}
	// 	var str = "let standards = ";
	// 	for(var i = 0; i < ws.pivotTables.length; ++i){
	// 		var res = getReportValues(ws.pivotTables[i]);
	// 		str += ws.pivotTables[i].asc_getName() + "\n";
	// 		str += "[\n";
	// 		for (var j = 0; j < res.length; ++j) {
	// 			str += JSON.stringify(res[j]);
	// 			if (j + 1 < res.length) {
	// 				str += ",\n";
	// 			} else {
	// 				str += "\n";
	// 			}
	// 		}
	// 		str += "]\n";
	// 	}
	// 	return str;
	// };
	// function getTestValuesMatrix(ws, range) {
	// 	if (!ws || !range) {
	// 		return "";
	// 	}
	// 	var res = getValues(ws, range);
	// 	var str = "[\n";
	// 	for (var i = 0; i < res.length; ++i) {
	// 		str += JSON.stringify(res[i]);
	// 		if (i + 1 < res.length) {
	// 			str += ",\n";
	// 		} else {
	// 			str += "\n";
	// 		}
	// 	}
	// 	str += "]\n";
	// 	return str;
	// };
	// let onDocumentContentReadyOld = AscCommon.baseEditorsApi.prototype.onDocumentContentReady;
	// AscCommon.baseEditorsApi.prototype.onDocumentContentReady = function() {
	// 	onDocumentContentReadyOld.call(this);
	// 	if(this.wbModel){
	// 		console.log('let testData = '+getTestValuesMatrix(this.wbModel.aWorksheets[1], AscCommonExcel.g_oRangeCache.getAscRange("B2:H3")));
	// 		console.log(getTestMatrix(this.wbModel.aWorksheets[0]));
	// 	}
	// };

	// To reproduce test in editor
	// 1) Open xlsx with pivot data in editor
	// 2) Execute preparation code in console
	// var api = Asc.editor;
	// var wb = api.wbModel;
	// var ws = wb.aWorksheets[0];
	// var pivotStyle = "PivotStyleDark23";
	// var reportRange = AscCommonExcel.g_oRangeCache.getAscRange("A3");
	// var dataRef = "Data!B2:H3";
	// 3) Execute part of code related to pivot in console. example fo 'testPivotMisc'
	// var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
	// pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
	// pivot.pivotTableDefinitionX14 = new Asc.CT_pivotTableDefinitionX14();
	// pivot.checkPivotFieldItems(0);
	// pivot.checkPivotFieldItems(1);
	// pivot.checkPivotFieldItems(2);
	//
	// var props = new Asc.CT_pivotTableDefinition();
	// props.asc_setName("new<&>pivot name");
	// props.asc_setTitle("Title");
	// props.asc_setDescription("Description");
	// pivot.asc_set(api, props);
	//
	// pivot.asc_addRowField(api, 0);
	// pivot.asc_addColField(api, 1);
	// pivot.asc_addColField(api, 2);
	// pivot.asc_addDataField(api, 5);
	// var props = new Asc.CT_pivotTableDefinition();
	// props.asc_setCompact(false);
	// props.asc_setOutline(true);
	// props.asc_setRowGrandTotals(false);
	// props.asc_setColGrandTotals(false);
	// pivot.asc_set(api, props);
	// var pivotField = pivot.asc_getPivotFields()[1];
	// props = new Asc.CT_PivotField();
	// props.asc_setDefaultSubtotal(false);
	// pivotField.asc_set(api, pivot, 1, props);

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

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function() {
			setTimeout(startTests, 0)
		}
	};
	window["Asc"]["editor"] = api;
var wb, ws, wsData, pivotStyle, tableName, defNameName, defNameLocalName, reportRange, testDataRange, testDataRange2,
	testDataRange3, testDataRange4, testDataRange5, testDataRange6, testDataRange7, testDataRangeHeader,
	testDataRangeTable,	testDataRefreshFieldSettings, testDataRangeFilters, testDataRangeNumFormat, testDataRangeDefName,
	testDataRangeDefNameLocal, testDataRefreshRecords, testDataRefreshStructure, testData, testData2, testData3,
	testData4, testData5, testData6, testData7, testDataFilter, testDataNumFormat, multiElem1, multiElem2, testDataHeader,
	testDataRecords, testDataStructure, addFormatTableOptions, defName, defNameLocal, dataRef, dataRef1Row, dataRef2,
	dataRef3, dataRef4, dataRef5, dataRef6, dataRef7, dataRefHeader, dataRefTable, dataRefTableColumn, dataRefDefName,
	dataRefDefNameLocal, dataRefFieldSettings, dataRefRecords, dataRefStructure, dataRefFilters, dataRefNumFormat;

	function openDocument(){
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api._openDocument(AscCommon.getEmpty());
		api._openOnClient();
		api.initCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		wb = api.wbModel;
		wb.handlers.add("getSelectionState", function() {
			return null;
		});
		ws = api.wbModel.aWorksheets[0];
		api.asc_insertWorksheet(["Data"]);
		wsData = wb.getWorksheetByName(["Data"], 0);
		api.asc_insertWorksheet(["Details"]);
		wsDetails = wb.getWorksheetByName(["Details"], 0);

		pivotStyle = "PivotStyleDark23";
		tableName = "Table1";
		defNameName = "defName";
		defNameLocalName = "defNameLocal";
		reportRange = AscCommonExcel.g_oRangeCache.getAscRange("A3");
		testDataRange = AscCommonExcel.g_oRangeCache.getAscRange("B2:H13");
		testDataRange2 = AscCommonExcel.g_oRangeCache.getAscRange("J2:P13");
		testDataRange3 = AscCommonExcel.g_oRangeCache.getAscRange("B15:H26");
		testDataRange4 = AscCommonExcel.g_oRangeCache.getAscRange("J15:P26");
		testDataRange5 = AscCommonExcel.g_oRangeCache.getAscRange("B28:H39");
		testDataRange6 = AscCommonExcel.g_oRangeCache.getAscRange("J28:P39");
		testDataRange7 = AscCommonExcel.g_oRangeCache.getAscRange("B41:H52");
		testDataRangeHeader = AscCommonExcel.g_oRangeCache.getAscRange("B54:O55");
		testDataRangeTable = AscCommonExcel.g_oRangeCache.getAscRange("B57:H68");
		testDataRefreshFieldSettings = AscCommonExcel.g_oRangeCache.getAscRange("B70:H81");
		testDataRangeFilters = AscCommonExcel.g_oRangeCache.getAscRange("B83:H93");
		testDataRangeNumFormat = AscCommonExcel.g_oRangeCache.getAscRange("B96:L107");
		testDataRangeDefName = AscCommonExcel.g_oRangeCache.getAscRange("J57:P68");
		testDataRangeDefNameLocal = AscCommonExcel.g_oRangeCache.getAscRange("R57:X68");
		testDataRefreshRecords = AscCommonExcel.g_oRangeCache.getAscRange("J70:P81");
		testDataRefreshStructure = AscCommonExcel.g_oRangeCache.getAscRange("R70:X81");
		testData = [
			["Region", "Gender", "Style", "Ship date", "Units", "Price", "Cost"],
			["East", "Boy", "Tee", "38383", "12", "11.04", "10.42"],
			["East", "Boy", "Golf", "38383", "12", "13", "12.6"],
			["East", "Boy", "Fancy", "38383", "12", "11.96", "11.74"],
			["East", "Girl", "Tee", "38383", "10", "11.27", "10.56"],
			["East", "Girl", "Golf", "38383", "10", "12.12", "11.95"],
			["East", "Girl", "Fancy", "38383", "10", "13.74", "13.33"],
			["West", "Boy", "Tee", "38383", "11", "11.44", "10.94"],
			["West", "Boy", "Golf", "38383", "11", "12.63", "11.73"],
			["West", "Boy", "Fancy", "38383", "11", "12.06", "11.51"],
			["West", "Girl", "Tee", "38383", "15", "13.42", "13.29"],
			["West", "Girl", "Golf", "38383", "15", "11.48", "10.67"]
		];
		testData2 = [
			["Region","Gender","Style","Ship date","Units","Price","Cost"],
			["East","Boy","Tee","38383","12","11.05","10.42"],
			["East","Boy","Golf","38383","12","","12.6"],
			["East","Boy","Fancy","38383","12","11.96","11.74"],
			["East","Girl","Tee","38383","10","","10.56"],
			["East","Girl","Golf","38383","10","12.12","11.95"],
			["East","Girl","Fancy","38383","10","","13.33"],
			["West","Boy","Tee","38383","11","11.44","10.94"],
			["West","Boy","Golf","38383","11","12.63","11.73"],
			["West","Boy","Fancy","38383","11","12.06","11.51"],
			["West","Girl","Tee","38383","15","13.42","13.29"],
			["West","Girl","Golf","38383","15","11.48","10.67"]
		];
		testData3 = [
			["Region","Gender","Style","Ship date","Units","Price","Cost"],
			["East","Boy","Tee","38383","12","11.05","10.42"],
			["East","Boy","Golf","38383","12","q","12.6"],
			["East","Boy","Fancy","38383","12","11.96","11.74"],
			["East","Girl","Tee","38383","10","w","10.56"],
			["East","Girl","Golf","38383","10","12.12","11.95"],
			["East","Girl","Fancy","38383","10","e","13.33"],
			["West","Boy","Tee","38383","11","11.44","10.94"],
			["West","Boy","Golf","38383","11","12.63","11.73"],
			["West","Boy","Fancy","38383","11","12.06","11.51"],
			["West","Girl","Tee","38383","15","13.42","13.29"],
			["West","Girl","Golf","38383","15","11.48","10.67"]
		];
		testData4 = [
			["Region","Gender","Style","Ship date","Units","Price","Cost"],
			["East","Boy","Tee","38383","12","11.05","10.42"],
			["East","Boy","Golf","38383","12","TRUE","12.6"],
			["East","Boy","Fancy","38383","12","11.96","11.74"],
			["East","Girl","Tee","38383","10","FALSE","10.56"],
			["East","Girl","Golf","38383","10","12.12","11.95"],
			["East","Girl","Fancy","38383","10","TRUE","13.33"],
			["West","Boy","Tee","38383","11","11.44","10.94"],
			["West","Boy","Golf","38383","11","12.63","11.73"],
			["West","Boy","Fancy","38383","11","12.06","11.51"],
			["West","Girl","Tee","38383","15","13.42","13.29"],
			["West","Girl","Golf","38383","15","11.48","10.67"]
		];
		testData5 = [
			["Region","Gender","Style","Ship date","Units","Price","Cost"],
			["East","Boy","Tee","38383","12","11.04","10.42"],
			["East","Boy","Golf","38383","12","#N/A","12.6"],
			["East","Boy","Fancy","38383","12","11.96","11.74"],
			["East","Girl","Tee","38383","10","#N/A","10.56"],
			["East","Girl","Golf","38383","10","12.12","11.95"],
			["East","Girl","Fancy","38383","10","#N/A","13.33"],
			["West","Boy","Tee","38383","11","11.44","10.94"],
			["West","Boy","Golf","38383","11","12.63","11.73"],
			["West","Boy","Fancy","38383","11","12.06","11.51"],
			["West","Girl","Tee","38383","15","13.42","13.29"],
			["West","Girl","Golf","38383","15","11.48","10.67"]
		];
		testData6 = [
			["Region","Gender","Style","Ship date","Units","Price","Cost"],
			["East","Boy","Tee","38383","12","11.04","10.42"],
			["East","Boy","Golf","38383","12",{format: "[$-409]m/d/yyyy h:mm AM/PM;@", value: new AscCommonExcel.CCellValue({number: 13})},"12.6"],
			["East","Boy","Fancy","38383","12","11.96","11.74"],
			["East","Girl","Tee","38383","10",{format: "[$-409]m/d/yyyy h:mm AM/PM;@", value: new AscCommonExcel.CCellValue({number: 11.27})},"10.56"],
			["East","Girl","Golf","38383","10","12.12","11.95"],
			["East","Girl","Fancy","38383","10",{format: "[$-409]m/d/yyyy h:mm AM/PM;@", value: new AscCommonExcel.CCellValue({number: 13.74})},"13.33"],
			["West","Boy","Tee","38383","11","11.44","10.94"],
			["West","Boy","Golf","38383","11","12.63","11.73"],
			["West","Boy","Fancy","38383","11","12.06","11.51"],
			["West","Girl","Tee","38383","15","13.42","13.29"],
			["West","Girl","Golf","38383","15","11.48","10.67"]
		];
		testData7 = [
			["Region","Gender","Style","Ship date","Units","Price","Cost"],
			["East","Boy","Tee","38383","12","","10.42"],
			["East","Boy","Golf","38383","12","","12.6"],
			["East","Boy","Fancy","38383","12","","11.74"],
			["East","Girl","Tee","38383","10","q","10.56"],
			["East","Girl","Golf","38383","10","","11.95"],
			["East","Girl","Fancy","38383","10","","13.33"],
			["West","Boy","Tee","38383","11","q","10.94"],
			["West","Boy","Golf","38383","11","1","11.73"],
			["West","Boy","Fancy","38383","11","","11.51"],
			["West","Girl","Tee","38383","15","2","13.29"],
			["West","Girl","Golf","38383","15","3","10.67"]
		];
		testDataFilter = [
			["Region","Gender","Style","Ship date","Units","Price","Cost"],
			["East","Boy","Tee","1","1","1","10.42"],
			["East","Boy","Golf","1","2","2","12.6"],
			["East","Boy","Fancy","2","2","3","11.74"],
			["East","Girl","Tee","2","3","4","10.56"],
			["East","Girl","Golf","3","3","5","11.95"],
			["East","Girl","Fancy","3","4","6","13.33"],
			["West","Boy","Tee","4","4","7","10.94"],
			["West","Boy","Golf","4","5","20","11.73"],
			["West","Boy","Fancy","5","5","20","11.51"],
			["West","Girl","Tee","6","6","20","13.29"]
		];
		testDataNumFormat = [
			["Text","Date","Units","Units2","Units3","Price","hasBlank","Mix","MixFormat","bool","error"],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Tee"})},  {format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38383})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 12})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 12})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 12})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 11.04})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10.42})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10.42})},                                     {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10.42})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 1})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#DIV/0!"})}],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Golf"})}, {format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38384})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 12})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 12})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 12})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 13})},   {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 12.6 })},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 12.6 })},                                     {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 12.6 })},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 1})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#DIV/0!"})}],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Fancy"})},{format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38385})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 12})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 12})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 12})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 11.96})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.74})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.74})},                                     {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.74})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 1})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#DIV/0!"})}],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Tee"})},  {format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38386})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 10})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 10})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 11.27})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue()},               {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "qwe"})}, {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.27})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 1})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#DIV/0!"})}],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Golf"})}, {format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38387})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 10})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 10})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 12.12})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.95})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.95})},                                     {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.95})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 0})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#DIV/0!"})}],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Fancy"})},{format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38388})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 10})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 10})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 13.74})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 13.33})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 13.33})},                                     {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 13.33})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 0})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#DIV/0!"})}],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Tee"})},  {format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38389})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 11})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 11})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 11.44})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10.94})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10.94})},                                     {format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10.94})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 0})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#DIV/0!"})}],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Golf"})}, {format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38390})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 11})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 11})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 12.63})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.73})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.73})},                                     {format: "0.00%",                                                             value: new AscCommonExcel.CCellValue({number: 11.73})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 0})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#N/A"})}   ],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Fancy"})},{format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38391})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 11})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 11})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 12.06})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.51})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 11.51})},                                     {format: "0.00%",                                                             value: new AscCommonExcel.CCellValue({number: 11.51})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 0})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#N/A"})}   ],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Tee"})},  {format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38392})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 15})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 15})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 15})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 13.42})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 13.29})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 13.29})},                                     {format: "0.00%",                                                             value: new AscCommonExcel.CCellValue({number: 13.29})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 0})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#N/A"})}   ],
			[{format: "\\q\\-@", value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.String, text: "Golf"})}, {format: "mm/dd/yy;@", value: new AscCommonExcel.CCellValue({number: 38393})},{format: "0.0%", value: new AscCommonExcel.CCellValue({number: 15})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 15})},{format: "0.0E+00", value: new AscCommonExcel.CCellValue({number: 15})},{format: '"$"#,##0.0', value: new AscCommonExcel.CCellValue({number: 11.48})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10.67})},{format: '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)', value: new AscCommonExcel.CCellValue({number: 10.67})},                                     {format: "0.00%",                                                             value: new AscCommonExcel.CCellValue({number: 10.67})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Bool, number: 0})},{value: new AscCommonExcel.CCellValue({type: AscCommon.CellValueType.Error, text: "#N/A"})}   ]
		];
		multiElem1 = new AscCommonExcel.CMultiTextElem();
		multiElem1.text = "qwe";
		multiElem2 = new AscCommonExcel.CMultiTextElem();
		multiElem2.text = "rty";
		testDataHeader = [
			["1","qwe","TRUE","#DIV/0!","1.234567891","12345678912", {format: "0.00", value: new AscCommonExcel.CCellValue({number: 1})},{format: "$#,##0.00", value: new AscCommonExcel.CCellValue({number: 2})},{format: "dddd\\, mmmm dd\\, yyyy", value: new AscCommonExcel.CCellValue({number: 3})},{format: "0.00%", value: new AscCommonExcel.CCellValue({number: 4})},{format: "0.00E+00", value: new AscCommonExcel.CCellValue({number: 5})},{value: new AscCommonExcel.CCellValue({multiText: [multiElem1, multiElem2]})},"1","qwe"],
			["1","1","1","1","1","1","1","1","1","1","1","1","1","1"]
		];
		testDataRecords = [
			["Region","Gender","Style","Ship date","Units","Price","Cost"],
			["East","Boy","Tee","38383","12","11.04","10.42"],
			["East","Boy","Golf","100000","12","13","12.6"],
			["East","Boy","Fancy","10","12","11.96","11.74"],
			["North","Girl","Tee","38383","10","11.27","10.56"],
			["East","Dog","Golf","38383","10","12.12","11.95"],
			["East","Girl","Fancy","38383","10","13.74","13.33"],
			["West","Boy","WWW","38383","11","11.44","10.94"],
			["West","Boy","Golf","38383","11","12.63","11.73"],
			["West","Boy","Fancy","38383","11","12.06","11.51"],
			["West","Girl","BBB","38383","15","13.42","13.29"],
			["West","Girl","Golf","38383","15","11.48","10.67"]
		];
		testDataStructure = [
			["NewField","Region","Style","NewUnits","Price","Gender","Cost"],
			["East","11.04","East","12","East","Boy","1"],
			["East","13","East","12","East","Boy","2"],
			["East","11.96","East","12","East","Boy","3"],
			["East","11.27","East","10","East","Girl","4"],
			["East","12.12","East","10","East","Girl","5"],
			["East","13.74","East","10","East","Girl","6"],
			["West","11.44","West","11","West","Boy","7"],
			["West","12.63","West","11","West","Boy","8"],
			["West","12.06","West","11","West","Boy","9"],
			["West","13.42","West","15","West","Girl","10"],
			["West","11.48","West","15","West","Girl","11"]
		];

		fillData(wsData, testData, testDataRange);
		fillData(wsData, testData2, testDataRange2);
		fillData(wsData, testData3, testDataRange3);
		fillData(wsData, testData4, testDataRange4);
		fillData(wsData, testData5, testDataRange5);
		fillData(wsData, testData6, testDataRange6);
		fillData(wsData, testData7, testDataRange7);
		fillData(wsData, testDataHeader, testDataRangeHeader);
		fillData(wsData, testData, testDataRangeTable);
		fillData(wsData, testDataFilter, testDataRangeFilters);
		fillData(wsData, testDataNumFormat, testDataRangeNumFormat);
		addFormatTableOptions = new AscCommonExcel.AddFormatTableOptions();
		addFormatTableOptions.asc_setRange(testDataRangeTable.getAbsName());
		addFormatTableOptions.asc_setIsTitle(true);
		wsData.autoFilters.addAutoFilter("TableStyleMedium2", testDataRangeTable, addFormatTableOptions);
		fillData(wsData, testData, testDataRangeDefName);
		defName = new Asc.asc_CDefName();
		defName.Name = defNameName;
		defName.Ref = wsData.getName() + "!" + testDataRangeDefName.getAbsName();
		api.asc_setDefinedNames(defName);
		fillData(wsData, testData, testDataRangeDefNameLocal);
		defNameLocal = new Asc.asc_CDefName();
		defNameLocal.Name = defNameLocalName;
		defNameLocal.Ref = wsData.getName() + "!" + testDataRangeDefNameLocal.getAbsName();
		defNameLocal.LocalSheetId = wsData.getId();
		api.asc_setDefinedNames(defNameLocal);
		fillData(wsData, testData, testDataRefreshFieldSettings);
		fillData(wsData, testDataRecords, testDataRefreshRecords);
		fillData(wsData, testDataStructure, testDataRefreshStructure);

		dataRef = wsData.getName() + "!" + testDataRange.getName();
		dataRef1Row = wsData.getName() + "!" + new Asc.Range(testDataRange.c1, testDataRange.r1, testDataRange.c2, testDataRange.r1 + 1).getName();
		dataRef2 = wsData.getName() + "!" + testDataRange2.getName();
		dataRef3 = wsData.getName() + "!" + testDataRange3.getName();
		dataRef4 = wsData.getName() + "!" + testDataRange4.getName();
		dataRef5 = wsData.getName() + "!" + testDataRange5.getName();
		dataRef6 = wsData.getName() + "!" + testDataRange6.getName();
		dataRef7 = wsData.getName() + "!" + testDataRange7.getName();
		dataRefHeader = wsData.getName() + "!" + testDataRangeHeader.getName();
		dataRefTable = tableName;
		dataRefTableColumn = tableName + '[[Gender]:[Price]]';
		dataRefDefName = defNameName;
		dataRefDefNameLocal = wsData.getName() + "!" + defNameLocalName;
		dataRefFieldSettings = wsData.getName() + "!" + testDataRefreshFieldSettings.getName();
		dataRefRecords = wsData.getName() + "!" + testDataRefreshRecords.getName();
		dataRefStructure = wsData.getName() + "!" + testDataRefreshStructure.getName();
		dataRefFilters = wsData.getName() + "!" + testDataRangeFilters.getName();
		dataRefNumFormat = wsData.getName() + "!" + testDataRangeNumFormat.getName();
	}


	function fillData(ws, data, range) {
		range = ws.getRange4(range.r1, range.c1);
		range.fillData(data);
	}
	function getReportValues(pivot) {
		var res = [];
		var range = new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange();
		pivot.GetWS().getRange3(range.r1, range.c1, range.r2, range.c2)._foreach(function(cell, r, c, r1, c1) {
			if (!res[r - r1]) {
				res[r - r1] = [];
			}
			res[r - r1][c - c1] = cell.getValue();
		});
		return res;
	}
	function checkReportValues(assert, pivot, standard, message) {
		let values = getReportValues(pivot);
		assert.deepEqual(values, standard, message);

		var isEmptyPivot = !(pivot.asc_getRowFields() || pivot.asc_getColumnFields() || pivot.asc_getDataFields() || pivot.asc_getPageFields());
		var styleError = "";
		var dataError = "";
		var ws = pivot.GetWS();
		ws._forEachCell(function(cell) {
			if (!testDataRange.contains(cell.nCol, cell.nRow)) {
				var inPivot = pivot.contains(cell.nCol, cell.nRow);
				var compiledStyle = ws.sheetMergedStyles.getStyle(ws.hiddenManager, cell.nRow, cell.nCol, ws);
				if (inPivot && !isEmptyPivot) {
					if (!compiledStyle || 0 === compiledStyle.table.length) {
						styleError = cell.getName();
					}
				} else {
					if (compiledStyle && 0 < compiledStyle.table.length) {
						styleError = cell.getName();
					}
					if (!cell.isEmptyTextString()) {
						dataError = cell.getName();
					}
				}
			}
		});
		assert.strictEqual(styleError, "", message + "_styleError");
		assert.strictEqual(dataError, "", message + "_dataError");
	}

	var memory = new AscCommon.CMemory();
	function Utf8ArrayToStr(array) {
		var out, i, len, c;
		var char2, char3;

		out = "";
		len = array.length;
		i = 0;
		while(i < len) {
			c = array[i++];
			switch(c >> 4)
			{
				case 0: case 1: case 2: case 3: case 4: case 5: case 6: case 7:
				// 0xxxxxxx
				out += String.fromCharCode(c);
				break;
				case 12: case 13:
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
	function getXml(pivot, addCacheDefinition){
		memory.Seek(0);
		pivot.toXml(memory);
		if(addCacheDefinition) {
			memory.WriteXmlString('\n\n');
			pivot.cacheDefinition.toXml(memory);
		}
		var buffer = new Uint8Array(memory.GetCurPosition());
		for (var i = 0; i < memory.GetCurPosition(); i++)
		{
			buffer[i] = memory.data[i];
		}
		if(typeof TextDecoder !== "undefined") {
			return new TextDecoder("utf-8").decode(buffer);
		} else {
			return Utf8ArrayToStr(buffer);
		}

	}

	function checkHistoryOperation(assert, pivot, standards, message, action) {
		let undoValues = getReportValues(pivot);
		return checkHistoryOperation2(assert, pivot, standards, message, undoValues, action, checkReportValues);
	}
	function checkHistoryOperation2(assert, pivot, standards, message, undoStandard, action, check, checkUndo) {
		var wb = pivot.GetWS().workbook;
		var xmlUndo = getXml(pivot, false);
		var pivotStart = pivot.clone();
		pivotStart.Id = pivot.Get_Id();

		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		action();
		AscCommon.History.EndTransaction();
		pivot = wb.getPivotTableById(pivot.Get_Id());
		check(assert, pivot, standards, message);
		var xmlDo = getXml(pivot, true);
		var changes = wb.SerializeHistory()[0];

		AscCommon.History.Undo();
		pivot = wb.getPivotTableById(pivot.Get_Id());
		check(assert, pivot, undoStandard, message + "_undo");
		assert.strictEqual(getXml(pivot, false), xmlUndo, message + "_undo_xml");

		AscCommon.History.Redo();
		pivot = wb.getPivotTableById(pivot.Get_Id());
		check(assert, pivot, standards, message + "_redo");
		assert.strictEqual(getXml(pivot, true), xmlDo, message + "_redo_xml");

		AscCommon.History.Undo();
		ws.deletePivotTable(pivot.Get_Id());
		pivot = pivotStart;
		ws.insertPivotTable(pivot, false, false);
		wb.DeserializeHistory(changes);
		pivot = wb.getPivotTableById(pivot.Get_Id());
		check(assert, pivot, standards, message + "_changes");
		assert.strictEqual(getXml(pivot, true), xmlDo, message + "_changes_xml");
		return pivot;
	}


	var standards = {
		"compact_0row_0col_0data": [
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""]
		],
		"compact_0row_0col_1data": [
				["Sum of Price"],
				["134.16"]
		],
		"compact_0row_0col_2data_col": [
				["Sum of Price", "Sum of Cost"],
				["134.16", "128.74"]
		],
		"compact_0row_0col_2data_row": [
				["Values", ""],
				["Sum of Price", "134.16"],
				["Sum of Cost", "128.74"]
		],
		"compact_0row_1col_0data": [
				["Column Labels", "", "", ""],
				["Fancy", "Golf", "Tee", "Grand Total"]
		],
		"compact_0row_1col_1data": [
				["", "Column Labels", "", "", ""],
				["", "Fancy", "Golf", "Tee", "Grand Total"],
				["Sum of Price", "37.76", "49.23", "47.17", "134.16"]
		],
		"compact_0row_1col_2data_col": [
				["Column Labels", "", "", "", "", "", "", ""],
				["Fancy", "", "Golf", "", "Tee", "", "Total Sum of Price", "Total Sum of Cost"],
				["Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "", ""],
				["37.76", "36.58", "49.23", "46.95", "47.17", "45.21", "134.16", "128.74"]
		],
		"compact_0row_1col_2data_row": [
				["", "Column Labels", "", "", ""],
				["Values", "Fancy", "Golf", "Tee", "Grand Total"],
				["Sum of Price", "37.76", "49.23", "47.17", "134.16"],
				["Sum of Cost", "36.58", "46.95", "45.21", "128.74"]
		],
		"compact_0row_2col_0data": [
				["Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""]
		],
		"compact_0row_2col_1data": [
				["", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["", "Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["", "10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""],
				["Sum of Price", "13.74", "12.06", "11.96", "37.76", "12.12", "12.63", "13", "11.48", "49.23", "11.27",
					"11.44", "11.04", "13.42", "47.17", "134.16"
				]
		],
		"compact_0row_2col_2data_col": [
				["Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
					"", "", "", "", "", "", "", ""
				],
				["Fancy", "", "", "", "", "", "Fancy Sum of Price", "Fancy Sum of Cost", "Golf", "", "", "", "", "", "",
					"", "Golf Sum of Price", "Golf Sum of Cost", "Tee", "", "", "", "", "", "", "", "Tee Sum of Price",
					"Tee Sum of Cost", "Total Sum of Price", "Total Sum of Cost"
				],
				["10", "", "11", "", "12", "", "", "", "10", "", "11", "", "12", "", "15", "", "", "", "10", "", "11",
					"", "12", "", "15", "", "", "", "", ""
				],
				["Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "", "",
					"Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost",
					"Sum of Price", "Sum of Cost", "", "", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost",
					"Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "", "", "", ""
				],
				["13.74", "13.33", "12.06", "11.51", "11.96", "11.74", "37.76", "36.58", "12.12", "11.95", "12.63",
					"11.73", "13", "12.6", "11.48", "10.67", "49.23", "46.95", "11.27", "10.56", "11.44", "10.94",
					"11.04", "10.42", "13.42", "13.29", "47.17", "45.21", "134.16", "128.74"
				]
		],
		"compact_0row_2col_2data_row": [
				["", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["", "Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["Values", "10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""],
				["Sum of Price", "13.74", "12.06", "11.96", "37.76", "12.12", "12.63", "13", "11.48", "49.23", "11.27",
					"11.44", "11.04", "13.42", "47.17", "134.16"
				],
				["Sum of Cost", "13.33", "11.51", "11.74", "36.58", "11.95", "11.73", "12.6", "10.67", "46.95", "10.56",
					"10.94", "10.42", "13.29", "45.21", "128.74"
				]
		],
		"compact_1row_0col_0data": [
				["Row Labels"],
				["East"],
				["West"],
				["Grand Total"]
		],
		"compact_1row_0col_1data": [
				["Row Labels", "Sum of Price"],
				["East", "73.13"],
				["West", "61.03"],
				["Grand Total", "134.16"]
		],
		"compact_1row_0col_2data_col": [
				["Row Labels", "Sum of Price", "Sum of Cost"],
				["East", "73.13", "70.6"],
				["West", "61.03", "58.14"],
				["Grand Total", "134.16", "128.74"]
		],
		"compact_1row_0col_2data_row": [
				["Row Labels", ""],
				["East", ""],
				["Sum of Price", "73.13"],
				["Sum of Cost", "70.6"],
				["West", ""],
				["Sum of Price", "61.03"],
				["Sum of Cost", "58.14"],
				["Total Sum of Price", "134.16"],
				["Total Sum of Cost", "128.74"]
		],
		"compact_1row_1col_0data": [
				["", "Column Labels", "", "", ""],
				["Row Labels", "Fancy", "Golf", "Tee", "Grand Total"],
				["East", "", "", "", ""],
				["West", "", "", "", ""],
				["Grand Total", "", "", "", ""]
		],
		"compact_1row_1col_1data": [
				["Sum of Price", "Column Labels", "", "", ""],
				["Row Labels", "Fancy", "Golf", "Tee", "Grand Total"],
				["East", "25.7", "25.12", "22.31", "73.13"],
				["West", "12.06", "24.11", "24.86", "61.03"],
				["Grand Total", "37.76", "49.23", "47.17", "134.16"]
		],
		"compact_1row_1col_2data_col": [
				["", "Column Labels", "", "", "", "", "", "", ""],
				["", "Fancy", "", "Golf", "", "Tee", "", "Total Sum of Price", "Total Sum of Cost"],
				["Row Labels", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price",
					"Sum of Cost", "", ""
				],
				["East", "25.7", "25.07", "25.12", "24.55", "22.31", "20.98", "73.13", "70.6"],
				["West", "12.06", "11.51", "24.11", "22.4", "24.86", "24.23", "61.03", "58.14"],
				["Grand Total", "37.76", "36.58", "49.23", "46.95", "47.17", "45.21", "134.16", "128.74"]
		],
		"compact_1row_1col_2data_row": [
				["", "Column Labels", "", "", ""],
				["Row Labels", "Fancy", "Golf", "Tee", "Grand Total"],
				["East", "", "", "", ""],
				["Sum of Price", "25.7", "25.12", "22.31", "73.13"],
				["Sum of Cost", "25.07", "24.55", "20.98", "70.6"],
				["West", "", "", "", ""],
				["Sum of Price", "12.06", "24.11", "24.86", "61.03"],
				["Sum of Cost", "11.51", "22.4", "24.23", "58.14"],
				["Total Sum of Price", "37.76", "49.23", "47.17", "134.16"],
				["Total Sum of Cost", "36.58", "46.95", "45.21", "128.74"]
		],
		"compact_1row_2col_0data": [
				["", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["", "Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["Row Labels", "10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""],
				["East", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["West", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Grand Total", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
		],
		"compact_1row_2col_1data": [
				["Sum of Price", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["", "Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["Row Labels", "10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""],
				["East", "13.74", "", "11.96", "25.7", "12.12", "", "13", "", "25.12", "11.27", "", "11.04", "",
					"22.31", "73.13"
				],
				["West", "", "12.06", "", "12.06", "", "12.63", "", "11.48", "24.11", "", "11.44", "", "13.42", "24.86",
					"61.03"
				],
				["Grand Total", "13.74", "12.06", "11.96", "37.76", "12.12", "12.63", "13", "11.48", "49.23", "11.27",
					"11.44", "11.04", "13.42", "47.17", "134.16"
				]
		],
		"compact_1row_2col_2data_col": [
				["", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
					"", "", "", "", "", "", "", "", ""
				],
				["", "Fancy", "", "", "", "", "", "Fancy Sum of Price", "Fancy Sum of Cost", "Golf", "", "", "", "", "",
					"", "", "Golf Sum of Price", "Golf Sum of Cost", "Tee", "", "", "", "", "", "", "",
					"Tee Sum of Price", "Tee Sum of Cost", "Total Sum of Price", "Total Sum of Cost"
				],
				["", "10", "", "11", "", "12", "", "", "", "10", "", "11", "", "12", "", "15", "", "", "", "10", "",
					"11", "", "12", "", "15", "", "", "", "", ""
				],
				["Row Labels", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price",
					"Sum of Cost", "", "", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price",
					"Sum of Cost", "Sum of Price", "Sum of Cost", "", "", "Sum of Price", "Sum of Cost", "Sum of Price",
					"Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "", "", "", ""
				],
				["East", "13.74", "13.33", "", "", "11.96", "11.74", "25.7", "25.07", "12.12", "11.95", "", "", "13",
					"12.6", "", "", "25.12", "24.55", "11.27", "10.56", "", "", "11.04", "10.42", "", "", "22.31",
					"20.98", "73.13", "70.6"
				],
				["West", "", "", "12.06", "11.51", "", "", "12.06", "11.51", "", "", "12.63", "11.73", "", "", "11.48",
					"10.67", "24.11", "22.4", "", "", "11.44", "10.94", "", "", "13.42", "13.29", "24.86", "24.23",
					"61.03", "58.14"
				],
				["Grand Total", "13.74", "13.33", "12.06", "11.51", "11.96", "11.74", "37.76", "36.58", "12.12",
					"11.95", "12.63", "11.73", "13", "12.6", "11.48", "10.67", "49.23", "46.95", "11.27", "10.56",
					"11.44", "10.94", "11.04", "10.42", "13.42", "13.29", "47.17", "45.21", "134.16", "128.74"
				]
		],
		"compact_1row_2col_2data_row": [
				["", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["", "Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["Row Labels", "10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""],
				["East", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Sum of Price", "13.74", "", "11.96", "25.7", "12.12", "", "13", "", "25.12", "11.27", "", "11.04", "",
					"22.31", "73.13"
				],
				["Sum of Cost", "13.33", "", "11.74", "25.07", "11.95", "", "12.6", "", "24.55", "10.56", "", "10.42",
					"", "20.98", "70.6"
				],
				["West", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Sum of Price", "", "12.06", "", "12.06", "", "12.63", "", "11.48", "24.11", "", "11.44", "", "13.42",
					"24.86", "61.03"
				],
				["Sum of Cost", "", "11.51", "", "11.51", "", "11.73", "", "10.67", "22.4", "", "10.94", "", "13.29",
					"24.23", "58.14"
				],
				["Total Sum of Price", "13.74", "12.06", "11.96", "37.76", "12.12", "12.63", "13", "11.48", "49.23",
					"11.27", "11.44", "11.04", "13.42", "47.17", "134.16"
				],
				["Total Sum of Cost", "13.33", "11.51", "11.74", "36.58", "11.95", "11.73", "12.6", "10.67", "46.95",
					"10.56", "10.94", "10.42", "13.29", "45.21", "128.74"
				]
		],
		"compact_2row_0col_0data": [
				["Row Labels"],
				["East"],
				["Boy"],
				["Girl"],
				["West"],
				["Boy"],
				["Girl"],
				["Grand Total"]
		],
		"compact_2row_0col_1data": [
				["Row Labels", "Sum of Price"],
				["East", "73.13"],
				["Boy", "36"],
				["Girl", "37.13"],
				["West", "61.03"],
				["Boy", "36.13"],
				["Girl", "24.9"],
				["Grand Total", "134.16"]
		],
		"compact_2row_0col_2data_col": [
				["Row Labels", "Sum of Price", "Sum of Cost"],
				["East", "73.13", "70.6"],
				["Boy", "36", "34.76"],
				["Girl", "37.13", "35.84"],
				["West", "61.03", "58.14"],
				["Boy", "36.13", "34.18"],
				["Girl", "24.9", "23.96"],
				["Grand Total", "134.16", "128.74"]
		],
		"compact_2row_0col_2data_row": [
				["Row Labels", ""],
				["East", ""],
				["Boy", ""],
				["Sum of Price", "36"],
				["Sum of Cost", "34.76"],
				["Girl", ""],
				["Sum of Price", "37.13"],
				["Sum of Cost", "35.84"],
				["East Sum of Price", "73.13"],
				["East Sum of Cost", "70.6"],
				["West", ""],
				["Boy", ""],
				["Sum of Price", "36.13"],
				["Sum of Cost", "34.18"],
				["Girl", ""],
				["Sum of Price", "24.9"],
				["Sum of Cost", "23.96"],
				["West Sum of Price", "61.03"],
				["West Sum of Cost", "58.14"],
				["Total Sum of Price", "134.16"],
				["Total Sum of Cost", "128.74"]
		],
		"compact_2row_1col_0data": [
				["", "Column Labels", "", "", ""],
				["Row Labels", "Fancy", "Golf", "Tee", "Grand Total"],
				["East", "", "", "", ""],
				["Boy", "", "", "", ""],
				["Girl", "", "", "", ""],
				["West", "", "", "", ""],
				["Boy", "", "", "", ""],
				["Girl", "", "", "", ""],
				["Grand Total", "", "", "", ""]
		],
		"compact_2row_1col_1data": [
				["Sum of Price", "Column Labels", "", "", ""],
				["Row Labels", "Fancy", "Golf", "Tee", "Grand Total"],
				["East", "25.7", "25.12", "22.31", "73.13"],
				["Boy", "11.96", "13", "11.04", "36"],
				["Girl", "13.74", "12.12", "11.27", "37.13"],
				["West", "12.06", "24.11", "24.86", "61.03"],
				["Boy", "12.06", "12.63", "11.44", "36.13"],
				["Girl", "", "11.48", "13.42", "24.9"],
				["Grand Total", "37.76", "49.23", "47.17", "134.16"]
		],
		"compact_2row_1col_2data_col": [
				["", "Column Labels", "", "", "", "", "", "", ""],
				["", "Fancy", "", "Golf", "", "Tee", "", "Total Sum of Price", "Total Sum of Cost"],
				["Row Labels", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price",
					"Sum of Cost", "", ""
				],
				["East", "25.7", "25.07", "25.12", "24.55", "22.31", "20.98", "73.13", "70.6"],
				["Boy", "11.96", "11.74", "13", "12.6", "11.04", "10.42", "36", "34.76"],
				["Girl", "13.74", "13.33", "12.12", "11.95", "11.27", "10.56", "37.13", "35.84"],
				["West", "12.06", "11.51", "24.11", "22.4", "24.86", "24.23", "61.03", "58.14"],
				["Boy", "12.06", "11.51", "12.63", "11.73", "11.44", "10.94", "36.13", "34.18"],
				["Girl", "", "", "11.48", "10.67", "13.42", "13.29", "24.9", "23.96"],
				["Grand Total", "37.76", "36.58", "49.23", "46.95", "47.17", "45.21", "134.16", "128.74"]
		],
		"compact_2row_1col_2data_row": [
				["", "Column Labels", "", "", ""],
				["Row Labels", "Fancy", "Golf", "Tee", "Grand Total"],
				["East", "", "", "", ""],
				["Boy", "", "", "", ""],
				["Sum of Price", "11.96", "13", "11.04", "36"],
				["Sum of Cost", "11.74", "12.6", "10.42", "34.76"],
				["Girl", "", "", "", ""],
				["Sum of Price", "13.74", "12.12", "11.27", "37.13"],
				["Sum of Cost", "13.33", "11.95", "10.56", "35.84"],
				["East Sum of Price", "25.7", "25.12", "22.31", "73.13"],
				["East Sum of Cost", "25.07", "24.55", "20.98", "70.6"],
				["West", "", "", "", ""],
				["Boy", "", "", "", ""],
				["Sum of Price", "12.06", "12.63", "11.44", "36.13"],
				["Sum of Cost", "11.51", "11.73", "10.94", "34.18"],
				["Girl", "", "", "", ""],
				["Sum of Price", "", "11.48", "13.42", "24.9"],
				["Sum of Cost", "", "10.67", "13.29", "23.96"],
				["West Sum of Price", "12.06", "24.11", "24.86", "61.03"],
				["West Sum of Cost", "11.51", "22.4", "24.23", "58.14"],
				["Total Sum of Price", "37.76", "49.23", "47.17", "134.16"],
				["Total Sum of Cost", "36.58", "46.95", "45.21", "128.74"]
		],
		"compact_2row_2col_0data": [
				["", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["", "Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["Row Labels", "10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""],
				["East", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Boy", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Girl", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["West", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Boy", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Girl", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Grand Total", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
		],
		"compact_2row_2col_1data": [
				["Sum of Price", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["", "Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["Row Labels", "10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""],
				["East", "13.74", "", "11.96", "25.7", "12.12", "", "13", "", "25.12", "11.27", "", "11.04", "",
					"22.31", "73.13"
				],
				["Boy", "", "", "11.96", "11.96", "", "", "13", "", "13", "", "", "11.04", "", "11.04", "36"],
				["Girl", "13.74", "", "", "13.74", "12.12", "", "", "", "12.12", "11.27", "", "", "", "11.27", "37.13"],
				["West", "", "12.06", "", "12.06", "", "12.63", "", "11.48", "24.11", "", "11.44", "", "13.42", "24.86",
					"61.03"
				],
				["Boy", "", "12.06", "", "12.06", "", "12.63", "", "", "12.63", "", "11.44", "", "", "11.44", "36.13"],
				["Girl", "", "", "", "", "", "", "", "11.48", "11.48", "", "", "", "13.42", "13.42", "24.9"],
				["Grand Total", "13.74", "12.06", "11.96", "37.76", "12.12", "12.63", "13", "11.48", "49.23", "11.27",
					"11.44", "11.04", "13.42", "47.17", "134.16"
				]
		],
		"compact_2row_2col_2data_col": [
				["", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
					"", "", "", "", "", "", "", "", ""
				],
				["", "Fancy", "", "", "", "", "", "Fancy Sum of Price", "Fancy Sum of Cost", "Golf", "", "", "", "", "",
					"", "", "Golf Sum of Price", "Golf Sum of Cost", "Tee", "", "", "", "", "", "", "",
					"Tee Sum of Price", "Tee Sum of Cost", "Total Sum of Price", "Total Sum of Cost"
				],
				["", "10", "", "11", "", "12", "", "", "", "10", "", "11", "", "12", "", "15", "", "", "", "10", "",
					"11", "", "12", "", "15", "", "", "", "", ""
				],
				["Row Labels", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price",
					"Sum of Cost", "", "", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price",
					"Sum of Cost", "Sum of Price", "Sum of Cost", "", "", "Sum of Price", "Sum of Cost", "Sum of Price",
					"Sum of Cost", "Sum of Price", "Sum of Cost", "Sum of Price", "Sum of Cost", "", "", "", ""
				],
				["East", "13.74", "13.33", "", "", "11.96", "11.74", "25.7", "25.07", "12.12", "11.95", "", "", "13",
					"12.6", "", "", "25.12", "24.55", "11.27", "10.56", "", "", "11.04", "10.42", "", "", "22.31",
					"20.98", "73.13", "70.6"
				],
				["Boy", "", "", "", "", "11.96", "11.74", "11.96", "11.74", "", "", "", "", "13", "12.6", "", "", "13",
					"12.6", "", "", "", "", "11.04", "10.42", "", "", "11.04", "10.42", "36", "34.76"
				],
				["Girl", "13.74", "13.33", "", "", "", "", "13.74", "13.33", "12.12", "11.95", "", "", "", "", "", "",
					"12.12", "11.95", "11.27", "10.56", "", "", "", "", "", "", "11.27", "10.56", "37.13", "35.84"
				],
				["West", "", "", "12.06", "11.51", "", "", "12.06", "11.51", "", "", "12.63", "11.73", "", "", "11.48",
					"10.67", "24.11", "22.4", "", "", "11.44", "10.94", "", "", "13.42", "13.29", "24.86", "24.23",
					"61.03", "58.14"
				],
				["Boy", "", "", "12.06", "11.51", "", "", "12.06", "11.51", "", "", "12.63", "11.73", "", "", "", "",
					"12.63", "11.73", "", "", "11.44", "10.94", "", "", "", "", "11.44", "10.94", "36.13", "34.18"
				],
				["Girl", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "11.48", "10.67", "11.48", "10.67", "",
					"", "", "", "", "", "13.42", "13.29", "13.42", "13.29", "24.9", "23.96"
				],
				["Grand Total", "13.74", "13.33", "12.06", "11.51", "11.96", "11.74", "37.76", "36.58", "12.12",
					"11.95", "12.63", "11.73", "13", "12.6", "11.48", "10.67", "49.23", "46.95", "11.27", "10.56",
					"11.44", "10.94", "11.04", "10.42", "13.42", "13.29", "47.17", "45.21", "134.16", "128.74"
				]
		],
		"compact_2row_2col_2data_col2": [
				["","Column Labels","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["","Sum of Price","","","Sum of Cost","","","","","Sum of Price","","","","Sum of Cost","","","","","","Sum of Price","","","","Sum of Cost","","","","","","",""],
				["Row Labels","10","11","12","10","11","12","","","10","11","12","15","10","11","12","15","","","10","11","12","15","10","11","12","15","","","",""],
				["East","13.74","","11.96","13.33","","11.74","25.7","25.07","12.12","","13","","11.95","","12.6","","25.12","24.55","11.27","","11.04","","10.56","","10.42","","22.31","20.98","73.13","70.6"],
				["Boy","","","11.96","","","11.74","11.96","11.74","","","13","","","","12.6","","13","12.6","","","11.04","","","","10.42","","11.04","10.42","36","34.76"],
				["Girl","13.74","","","13.33","","","13.74","13.33","12.12","","","","11.95","","","","12.12","11.95","11.27","","","","10.56","","","","11.27","10.56","37.13","35.84"],
				["West","","12.06","","","11.51","","12.06","11.51","","12.63","","11.48","","11.73","","10.67","24.11","22.4","","11.44","","13.42","","10.94","","13.29","24.86","24.23","61.03","58.14"],
				["Boy","","12.06","","","11.51","","12.06","11.51","","12.63","","","","11.73","","","12.63","11.73","","11.44","","","","10.94","","","11.44","10.94","36.13","34.18"],
				["Girl","","","","","","","","","","","","11.48","","","","10.67","11.48","10.67","","","","13.42","","","","13.29","13.42","13.29","24.9","23.96"],
				["Grand Total","13.74","12.06","11.96","13.33","11.51","11.74","37.76","36.58","12.12","12.63","13","11.48","11.95","11.73","12.6","10.67","49.23","46.95","11.27","11.44","11.04","13.42","10.56","10.94","10.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"compact_2row_2col_2data_col1": [
				["","Column Labels","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","Sum of Price","","","","","","","","","","","","","","Sum of Cost","","","","","","","","","","","","","","Total Sum of Price","Total Sum of Cost"],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","",""],
				["Row Labels","10","11","12","","10","11","12","15","","10","11","12","15","","10","11","12","","10","11","12","15","","10","11","12","15","","",""],
				["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","73.13","70.6"],
				["Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","36","34.76"],
				["Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","37.13","35.84"],
				["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","61.03","58.14"],
				["Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","36.13","34.18"],
				["Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","","","","","","","","10.67","10.67","","","","13.29","13.29","24.9","23.96"],
				["Grand Total","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","134.16","128.74"]
		],
		"compact_2row_2col_2data_row": [
				["", "Column Labels", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["", "Fancy", "", "", "Fancy Total", "Golf", "", "", "", "Golf Total", "Tee", "", "", "", "Tee Total",
					"Grand Total"
				],
				["Row Labels", "10", "11", "12", "", "10", "11", "12", "15", "", "10", "11", "12", "15", "", ""],
				["East", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Boy", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Sum of Price", "", "", "11.96", "11.96", "", "", "13", "", "13", "", "", "11.04", "", "11.04", "36"],
				["Sum of Cost", "", "", "11.74", "11.74", "", "", "12.6", "", "12.6", "", "", "10.42", "", "10.42",
					"34.76"
				],
				["Girl", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Sum of Price", "13.74", "", "", "13.74", "12.12", "", "", "", "12.12", "11.27", "", "", "", "11.27",
					"37.13"
				],
				["Sum of Cost", "13.33", "", "", "13.33", "11.95", "", "", "", "11.95", "10.56", "", "", "", "10.56",
					"35.84"
				],
				["East Sum of Price", "13.74", "", "11.96", "25.7", "12.12", "", "13", "", "25.12", "11.27", "",
					"11.04", "", "22.31", "73.13"
				],
				["East Sum of Cost", "13.33", "", "11.74", "25.07", "11.95", "", "12.6", "", "24.55", "10.56", "",
					"10.42", "", "20.98", "70.6"
				],
				["West", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Boy", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Sum of Price", "", "12.06", "", "12.06", "", "12.63", "", "", "12.63", "", "11.44", "", "", "11.44",
					"36.13"
				],
				["Sum of Cost", "", "11.51", "", "11.51", "", "11.73", "", "", "11.73", "", "10.94", "", "", "10.94",
					"34.18"
				],
				["Girl", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
				["Sum of Price", "", "", "", "", "", "", "", "11.48", "11.48", "", "", "", "13.42", "13.42", "24.9"],
				["Sum of Cost", "", "", "", "", "", "", "", "10.67", "10.67", "", "", "", "13.29", "13.29", "23.96"],
				["West Sum of Price", "", "12.06", "", "12.06", "", "12.63", "", "11.48", "24.11", "", "11.44", "",
					"13.42", "24.86", "61.03"
				],
				["West Sum of Cost", "", "11.51", "", "11.51", "", "11.73", "", "10.67", "22.4", "", "10.94", "",
					"13.29", "24.23", "58.14"
				],
				["Total Sum of Price", "13.74", "12.06", "11.96", "37.76", "12.12", "12.63", "13", "11.48", "49.23",
					"11.27", "11.44", "11.04", "13.42", "47.17", "134.16"
				],
				["Total Sum of Cost", "13.33", "11.51", "11.74", "36.58", "11.95", "11.73", "12.6", "10.67", "46.95",
					"10.56", "10.94", "10.42", "13.29", "45.21", "128.74"
				]
		],
		"compact_2row_2col_2data_row2": [
				["","Column Labels","","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Row Labels","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","","","","","","","","","","","","","",""],
				["Sum of Price","","","","","","","","","","","","","","",""],
				["Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["Sum of Cost","","","","","","","","","","","","","","",""],
				["Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["East Sum of Price","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["East Sum of Cost","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["West","","","","","","","","","","","","","","",""],
				["Sum of Price","","","","","","","","","","","","","","",""],
				["Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["Sum of Cost","","","","","","","","","","","","","","",""],
				["Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["West Sum of Price","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["West Sum of Cost","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"compact_2row_2col_2data_row1": [
				["","Column Labels","","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Row Labels","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["Sum of Price","","","","","","","","","","","","","","",""],
				["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["Sum of Cost","","","","","","","","","","","","","","",""],
				["East","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["West","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["Total Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"outline_0row_0col_0data": [
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""],
				["", "", ""]
		],
		"outline_0row_0col_1data": [
				["Sum of Price"],
				["134.16"]
		],
		"outline_0row_0col_2data_col": [
				["Sum of Price","Sum of Cost"],
				["134.16","128.74"]
		],
		"outline_0row_0col_2data_row": [
				["Values",""],
				["Sum of Price","134.16"],
				["Sum of Cost","128.74"]
		],
		"outline_0row_1col_0data": [
				["Style","","",""],
				["Fancy","Golf","Tee","Grand Total"]
		],
		"outline_0row_1col_1data": [
				["","Style","","",""],
				["","Fancy","Golf","Tee","Grand Total"],
				["Sum of Price","37.76","49.23","47.17","134.16"]
		],
		"outline_0row_1col_2data_col": [
				["Style","Values","","","","","",""],
				["Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
				["Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
				["37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"outline_0row_1col_2data_row": [
				["","Style","","",""],
				["Values","Fancy","Golf","Tee","Grand Total"],
				["Sum of Price","37.76","49.23","47.17","134.16"],
				["Sum of Cost","36.58","46.95","45.21","128.74"]
		],
		"outline_0row_2col_0data": [
				["Style","Units","","","","","","","","","","","","",""],
				["Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["10","11","12","","10","11","12","15","","10","11","12","15","",""]
		],
		"outline_0row_2col_1data": [
				["","Style","Units","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"outline_0row_2col_2data_col": [
				["Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
				["Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
				["13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"outline_0row_2col_2data_row": [
				["","Style","Units","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"outline_1row_0col_0data": [
				["Region"],
				["East"],
				["West"],
				["Grand Total"]
		],
		"outline_1row_0col_1data": [
				["Region","Sum of Price"],
				["East","73.13"],
				["West","61.03"],
				["Grand Total","134.16"]
		],
		"outline_1row_0col_2data_col": [
				["Region","Sum of Price","Sum of Cost"],
				["East","73.13","70.6"],
				["West","61.03","58.14"],
				["Grand Total","134.16","128.74"]
		],
		"outline_1row_0col_2data_row": [
				["Region","Values",""],
				["East","",""],
				["","Sum of Price","73.13"],
				["","Sum of Cost","70.6"],
				["West","",""],
				["","Sum of Price","61.03"],
				["","Sum of Cost","58.14"],
				["Total Sum of Price","","134.16"],
				["Total Sum of Cost","","128.74"]
		],
		"outline_1row_1col_0data": [
				["","Style","","",""],
				["Region","Fancy","Golf","Tee","Grand Total"],
				["East","","","",""],
				["West","","","",""],
				["Grand Total","","","",""]
		],
		"outline_1row_1col_1data": [
				["Sum of Price","Style","","",""],
				["Region","Fancy","Golf","Tee","Grand Total"],
				["East","25.7","25.12","22.31","73.13"],
				["West","12.06","24.11","24.86","61.03"],
				["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"outline_1row_1col_2data_col": [
				["","Style","Values","","","","","",""],
				["","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
				["Region","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
				["East","25.7","25.07","25.12","24.55","22.31","20.98","73.13","70.6"],
				["West","12.06","11.51","24.11","22.4","24.86","24.23","61.03","58.14"],
				["Grand Total","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"outline_1row_1col_2data_row": [
				["","","Style","","",""],
				["Region","Values","Fancy","Golf","Tee","Grand Total"],
				["East","","","","",""],
				["","Sum of Price","25.7","25.12","22.31","73.13"],
				["","Sum of Cost","25.07","24.55","20.98","70.6"],
				["West","","","","",""],
				["","Sum of Price","12.06","24.11","24.86","61.03"],
				["","Sum of Cost","11.51","22.4","24.23","58.14"],
				["Total Sum of Price","","37.76","49.23","47.17","134.16"],
				["Total Sum of Cost","","36.58","46.95","45.21","128.74"]
		],
		"outline_1row_2col_0data": [
				["","Style","Units","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","","","","","","","","","","","","","",""],
				["West","","","","","","","","","","","","","","",""],
				["Grand Total","","","","","","","","","","","","","","",""]
		],
		"outline_1row_2col_1data": [
				["Sum of Price","Style","Units","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["Grand Total","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"outline_1row_2col_2data_col": [
				["","Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
				["Region","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
				["East","13.74","13.33","","","11.96","11.74","25.7","25.07","12.12","11.95","","","13","12.6","","","25.12","24.55","11.27","10.56","","","11.04","10.42","","","22.31","20.98","73.13","70.6"],
				["West","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","11.48","10.67","24.11","22.4","","","11.44","10.94","","","13.42","13.29","24.86","24.23","61.03","58.14"],
				["Grand Total","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"outline_1row_2col_2data_row": [
				["","","Style","Units","","","","","","","","","","","","",""],
				["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","","","","","","","","","","","","","","",""],
				["","Sum of Price","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["","Sum of Cost","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["West","","","","","","","","","","","","","","","",""],
				["","Sum of Price","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["","Sum of Cost","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"outline_2row_0col_0data": [
				["Region","Gender"],
				["East",""],
				["","Boy"],
				["","Girl"],
				["West",""],
				["","Boy"],
				["","Girl"],
				["Grand Total",""]
		],
		"outline_2row_0col_1data": [
				["Region","Gender","Sum of Price"],
				["East","","73.13"],
				["","Boy","36"],
				["","Girl","37.13"],
				["West","","61.03"],
				["","Boy","36.13"],
				["","Girl","24.9"],
				["Grand Total","","134.16"]
		],
		"outline_2row_0col_2data_col": [
				["Region","Gender","Sum of Price","Sum of Cost"],
				["East","","73.13","70.6"],
				["","Boy","36","34.76"],
				["","Girl","37.13","35.84"],
				["West","","61.03","58.14"],
				["","Boy","36.13","34.18"],
				["","Girl","24.9","23.96"],
				["Grand Total","","134.16","128.74"]
		],
		"outline_2row_0col_2data_row": [
				["Region","Gender","Values",""],
				["East","","",""],
				["","Boy","",""],
				["","","Sum of Price","36"],
				["","","Sum of Cost","34.76"],
				["","Girl","",""],
				["","","Sum of Price","37.13"],
				["","","Sum of Cost","35.84"],
				["East Sum of Price","","","73.13"],
				["East Sum of Cost","","","70.6"],
				["West","","",""],
				["","Boy","",""],
				["","","Sum of Price","36.13"],
				["","","Sum of Cost","34.18"],
				["","Girl","",""],
				["","","Sum of Price","24.9"],
				["","","Sum of Cost","23.96"],
				["West Sum of Price","","","61.03"],
				["West Sum of Cost","","","58.14"],
				["Total Sum of Price","","","134.16"],
				["Total Sum of Cost","","","128.74"]
		],
		"outline_2row_1col_0data": [
				["","","Style","","",""],
				["Region","Gender","Fancy","Golf","Tee","Grand Total"],
				["East","","","","",""],
				["","Boy","","","",""],
				["","Girl","","","",""],
				["West","","","","",""],
				["","Boy","","","",""],
				["","Girl","","","",""],
				["Grand Total","","","","",""]
		],
		"outline_2row_1col_1data": [
				["Sum of Price","","Style","","",""],
				["Region","Gender","Fancy","Golf","Tee","Grand Total"],
				["East","","25.7","25.12","22.31","73.13"],
				["","Boy","11.96","13","11.04","36"],
				["","Girl","13.74","12.12","11.27","37.13"],
				["West","","12.06","24.11","24.86","61.03"],
				["","Boy","12.06","12.63","11.44","36.13"],
				["","Girl","","11.48","13.42","24.9"],
				["Grand Total","","37.76","49.23","47.17","134.16"]
		],
		"outline_2row_1col_2data_col": [
				["","","Style","Values","","","","","",""],
				["","","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
				["Region","Gender","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
				["East","","25.7","25.07","25.12","24.55","22.31","20.98","73.13","70.6"],
				["","Boy","11.96","11.74","13","12.6","11.04","10.42","36","34.76"],
				["","Girl","13.74","13.33","12.12","11.95","11.27","10.56","37.13","35.84"],
				["West","","12.06","11.51","24.11","22.4","24.86","24.23","61.03","58.14"],
				["","Boy","12.06","11.51","12.63","11.73","11.44","10.94","36.13","34.18"],
				["","Girl","","","11.48","10.67","13.42","13.29","24.9","23.96"],
				["Grand Total","","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"outline_2row_1col_2data_row": [
				["","","","Style","","",""],
				["Region","Gender","Values","Fancy","Golf","Tee","Grand Total"],
				["East","","","","","",""],
				["","Boy","","","","",""],
				["","","Sum of Price","11.96","13","11.04","36"],
				["","","Sum of Cost","11.74","12.6","10.42","34.76"],
				["","Girl","","","","",""],
				["","","Sum of Price","13.74","12.12","11.27","37.13"],
				["","","Sum of Cost","13.33","11.95","10.56","35.84"],
				["East Sum of Price","","","25.7","25.12","22.31","73.13"],
				["East Sum of Cost","","","25.07","24.55","20.98","70.6"],
				["West","","","","","",""],
				["","Boy","","","","",""],
				["","","Sum of Price","12.06","12.63","11.44","36.13"],
				["","","Sum of Cost","11.51","11.73","10.94","34.18"],
				["","Girl","","","","",""],
				["","","Sum of Price","","11.48","13.42","24.9"],
				["","","Sum of Cost","","10.67","13.29","23.96"],
				["West Sum of Price","","","12.06","24.11","24.86","61.03"],
				["West Sum of Cost","","","11.51","22.4","24.23","58.14"],
				["Total Sum of Price","","","37.76","49.23","47.17","134.16"],
				["Total Sum of Cost","","","36.58","46.95","45.21","128.74"]
		],
		"outline_2row_2col_0data": [
				["","","Style","Units","","","","","","","","","","","","",""],
				["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","","","","","","","","","","","","","","",""],
				["","Boy","","","","","","","","","","","","","","",""],
				["","Girl","","","","","","","","","","","","","","",""],
				["West","","","","","","","","","","","","","","","",""],
				["","Boy","","","","","","","","","","","","","","",""],
				["","Girl","","","","","","","","","","","","","","",""],
				["Grand Total","","","","","","","","","","","","","","","",""]
		],
		"outline_2row_2col_1data": [
				["Sum of Price","","Style","Units","","","","","","","","","","","","",""],
				["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["West","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["Grand Total","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"outline_2row_2col_2data_col": [
				["","","Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["","","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
				["Region","Gender","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
				["East","","13.74","13.33","","","11.96","11.74","25.7","25.07","12.12","11.95","","","13","12.6","","","25.12","24.55","11.27","10.56","","","11.04","10.42","","","22.31","20.98","73.13","70.6"],
				["","Boy","","","","","11.96","11.74","11.96","11.74","","","","","13","12.6","","","13","12.6","","","","","11.04","10.42","","","11.04","10.42","36","34.76"],
				["","Girl","13.74","13.33","","","","","13.74","13.33","12.12","11.95","","","","","","","12.12","11.95","11.27","10.56","","","","","","","11.27","10.56","37.13","35.84"],
				["West","","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","11.48","10.67","24.11","22.4","","","11.44","10.94","","","13.42","13.29","24.86","24.23","61.03","58.14"],
				["","Boy","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","","","12.63","11.73","","","11.44","10.94","","","","","11.44","10.94","36.13","34.18"],
				["","Girl","","","","","","","","","","","","","","","11.48","10.67","11.48","10.67","","","","","","","13.42","13.29","13.42","13.29","24.9","23.96"],
				["Grand Total","","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"outline_2row_2col_2data_col2": [
				["","","Style","Values","Units","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["","","Sum of Price","","","Sum of Cost","","","","","Sum of Price","","","","Sum of Cost","","","","","","Sum of Price","","","","Sum of Cost","","","","","","",""],
				["Region","Gender","10","11","12","10","11","12","","","10","11","12","15","10","11","12","15","","","10","11","12","15","10","11","12","15","","","",""],
				["East","","13.74","","11.96","13.33","","11.74","25.7","25.07","12.12","","13","","11.95","","12.6","","25.12","24.55","11.27","","11.04","","10.56","","10.42","","22.31","20.98","73.13","70.6"],
				["","Boy","","","11.96","","","11.74","11.96","11.74","","","13","","","","12.6","","13","12.6","","","11.04","","","","10.42","","11.04","10.42","36","34.76"],
				["","Girl","13.74","","","13.33","","","13.74","13.33","12.12","","","","11.95","","","","12.12","11.95","11.27","","","","10.56","","","","11.27","10.56","37.13","35.84"],
				["West","","","12.06","","","11.51","","12.06","11.51","","12.63","","11.48","","11.73","","10.67","24.11","22.4","","11.44","","13.42","","10.94","","13.29","24.86","24.23","61.03","58.14"],
				["","Boy","","12.06","","","11.51","","12.06","11.51","","12.63","","","","11.73","","","12.63","11.73","","11.44","","","","10.94","","","11.44","10.94","36.13","34.18"],
				["","Girl","","","","","","","","","","","","11.48","","","","10.67","11.48","10.67","","","","13.42","","","","13.29","13.42","13.29","24.9","23.96"],
				["Grand Total","","13.74","12.06","11.96","13.33","11.51","11.74","37.76","36.58","12.12","12.63","13","11.48","11.95","11.73","12.6","10.67","49.23","46.95","11.27","11.44","11.04","13.42","10.56","10.94","10.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"outline_2row_2col_2data_col1": [
				["","","Values","Style","Units","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","","Sum of Price","","","","","","","","","","","","","","Sum of Cost","","","","","","","","","","","","","","Total Sum of Price","Total Sum of Cost"],
				["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","",""],
				["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","","10","11","12","","10","11","12","15","","10","11","12","15","","",""],
				["East","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","73.13","70.6"],
				["","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","36","34.76"],
				["","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","37.13","35.84"],
				["West","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","61.03","58.14"],
				["","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","36.13","34.18"],
				["","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","","","","","","","","10.67","10.67","","","","13.29","13.29","24.9","23.96"],
				["Grand Total","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","134.16","128.74"]
		],
		"outline_2row_2col_2data_row": [
				["","","","Style","Units","","","","","","","","","","","","",""],
				["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Gender","Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","","","","","","","","","","","","","","","",""],
				["","Boy","","","","","","","","","","","","","","","",""],
				["","","Sum of Price","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["","","Sum of Cost","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["","Girl","","","","","","","","","","","","","","","",""],
				["","","Sum of Price","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["","","Sum of Cost","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["East Sum of Price","","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["East Sum of Cost","","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["West","","","","","","","","","","","","","","","","",""],
				["","Boy","","","","","","","","","","","","","","","",""],
				["","","Sum of Price","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["","","Sum of Cost","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["","Girl","","","","","","","","","","","","","","","",""],
				["","","Sum of Price","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["","","Sum of Cost","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["West Sum of Price","","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["West Sum of Cost","","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"outline_2row_2col_2data_row2": [
				["","","","Style","Units","","","","","","","","","","","","",""],
				["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Values","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","","","","","","","","","","","","","","","",""],
				["","Sum of Price","","","","","","","","","","","","","","","",""],
				["","","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["","","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["","Sum of Cost","","","","","","","","","","","","","","","",""],
				["","","Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["","","Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["East Sum of Price","","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["East Sum of Cost","","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["West","","","","","","","","","","","","","","","","",""],
				["","Sum of Price","","","","","","","","","","","","","","","",""],
				["","","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["","","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["","Sum of Cost","","","","","","","","","","","","","","","",""],
				["","","Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["","","Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["West Sum of Price","","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["West Sum of Cost","","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"outline_2row_2col_2data_row1": [
				["","","","Style","Units","","","","","","","","","","","","",""],
				["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Values","Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["Sum of Price","","","","","","","","","","","","","","","","",""],
				["","East","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["","","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["","","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["","West","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["","","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["","","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["Sum of Cost","","","","","","","","","","","","","","","","",""],
				["","East","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["","","Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["","","Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["","West","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["","","Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["","","Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"tabular_0row_0col_0data": [
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""],
				["","",""]
		],
		"tabular_0row_0col_1data": [
				["Sum of Price"],
				["134.16"]
		],
		"tabular_0row_0col_2data_col": [
				["Sum of Price","Sum of Cost"],
				["134.16","128.74"]
		],
		"tabular_0row_0col_2data_row": [
				["Values",""],
				["Sum of Price","134.16"],
				["Sum of Cost","128.74"]
		],
		"tabular_0row_1col_0data": [
				["Style","","",""],
				["Fancy","Golf","Tee","Grand Total"]
		],
		"tabular_0row_1col_1data": [
				["","Style","","",""],
				["","Fancy","Golf","Tee","Grand Total"],
				["Sum of Price","37.76","49.23","47.17","134.16"]
		],
		"tabular_0row_1col_2data_col": [
				["Style","Values","","","","","",""],
				["Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
				["Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
				["37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"tabular_0row_1col_2data_row": [
				["","Style","","",""],
				["Values","Fancy","Golf","Tee","Grand Total"],
				["Sum of Price","37.76","49.23","47.17","134.16"],
				["Sum of Cost","36.58","46.95","45.21","128.74"]
		],
		"tabular_0row_2col_0data": [
				["Style","Units","","","","","","","","","","","","",""],
				["Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["10","11","12","","10","11","12","15","","10","11","12","15","",""]
		],
		"tabular_0row_2col_1data": [
				["","Style","Units","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"tabular_0row_2col_2data_col": [
				["Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
				["Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
				["13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"tabular_0row_2col_2data_row": [
				["","Style","Units","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"tabular_1row_0col_0data": [
				["Region"],
				["East"],
				["West"],
				["Grand Total"]
		],
		"tabular_1row_0col_1data": [
				["Region","Sum of Price"],
				["East","73.13"],
				["West","61.03"],
				["Grand Total","134.16"]
		],
		"tabular_1row_0col_2data_col": [
				["Region","Sum of Price","Sum of Cost"],
				["East","73.13","70.6"],
				["West","61.03","58.14"],
				["Grand Total","134.16","128.74"]
		],
		"tabular_1row_0col_2data_row": [
				["Region","Values",""],
				["East","Sum of Price","73.13"],
				["","Sum of Cost","70.6"],
				["West","Sum of Price","61.03"],
				["","Sum of Cost","58.14"],
				["Total Sum of Price","","134.16"],
				["Total Sum of Cost","","128.74"]
		],
		"tabular_1row_1col_0data": [
				["","Style","","",""],
				["Region","Fancy","Golf","Tee","Grand Total"],
				["East","","","",""],
				["West","","","",""],
				["Grand Total","","","",""]
		],
		"tabular_1row_1col_1data": [
				["Sum of Price","Style","","",""],
				["Region","Fancy","Golf","Tee","Grand Total"],
				["East","25.7","25.12","22.31","73.13"],
				["West","12.06","24.11","24.86","61.03"],
				["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"tabular_1row_1col_2data_col": [
				["","Style","Values","","","","","",""],
				["","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
				["Region","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
				["East","25.7","25.07","25.12","24.55","22.31","20.98","73.13","70.6"],
				["West","12.06","11.51","24.11","22.4","24.86","24.23","61.03","58.14"],
				["Grand Total","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"tabular_1row_1col_2data_row": [
				["","","Style","","",""],
				["Region","Values","Fancy","Golf","Tee","Grand Total"],
				["East","Sum of Price","25.7","25.12","22.31","73.13"],
				["","Sum of Cost","25.07","24.55","20.98","70.6"],
				["West","Sum of Price","12.06","24.11","24.86","61.03"],
				["","Sum of Cost","11.51","22.4","24.23","58.14"],
				["Total Sum of Price","","37.76","49.23","47.17","134.16"],
				["Total Sum of Cost","","36.58","46.95","45.21","128.74"]
		],
		"tabular_1row_2col_0data": [
				["","Style","Units","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","","","","","","","","","","","","","",""],
				["West","","","","","","","","","","","","","","",""],
				["Grand Total","","","","","","","","","","","","","","",""]
		],
		"tabular_1row_2col_1data": [
				["Sum of Price","Style","Units","","","","","","","","","","","","",""],
				["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["Grand Total","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"tabular_1row_2col_2data_col": [
				["","Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
				["Region","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
				["East","13.74","13.33","","","11.96","11.74","25.7","25.07","12.12","11.95","","","13","12.6","","","25.12","24.55","11.27","10.56","","","11.04","10.42","","","22.31","20.98","73.13","70.6"],
				["West","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","11.48","10.67","24.11","22.4","","","11.44","10.94","","","13.42","13.29","24.86","24.23","61.03","58.14"],
				["Grand Total","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"tabular_1row_2col_2data_row": [
				["","","Style","Units","","","","","","","","","","","","",""],
				["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","Sum of Price","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["","Sum of Cost","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["West","Sum of Price","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["","Sum of Cost","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"tabular_2row_0col_0data": [
				["Region","Gender"],
				["East","Boy"],
				["","Girl"],
				["East Total",""],
				["West","Boy"],
				["","Girl"],
				["West Total",""],
				["Grand Total",""]
		],
		"tabular_2row_0col_1data": [
				["Region","Gender","Sum of Price"],
				["East","Boy","36"],
				["","Girl","37.13"],
				["East Total","","73.13"],
				["West","Boy","36.13"],
				["","Girl","24.9"],
				["West Total","","61.03"],
				["Grand Total","","134.16"]
		],
		"tabular_2row_0col_2data_col": [
				["Region","Gender","Sum of Price","Sum of Cost"],
				["East","Boy","36","34.76"],
				["","Girl","37.13","35.84"],
				["East Total","","73.13","70.6"],
				["West","Boy","36.13","34.18"],
				["","Girl","24.9","23.96"],
				["West Total","","61.03","58.14"],
				["Grand Total","","134.16","128.74"]
		],
		"tabular_2row_0col_2data_row": [
				["Region","Gender","Values",""],
				["East","Boy","Sum of Price","36"],
				["","","Sum of Cost","34.76"],
				["","Girl","Sum of Price","37.13"],
				["","","Sum of Cost","35.84"],
				["East Sum of Price","","","73.13"],
				["East Sum of Cost","","","70.6"],
				["West","Boy","Sum of Price","36.13"],
				["","","Sum of Cost","34.18"],
				["","Girl","Sum of Price","24.9"],
				["","","Sum of Cost","23.96"],
				["West Sum of Price","","","61.03"],
				["West Sum of Cost","","","58.14"],
				["Total Sum of Price","","","134.16"],
				["Total Sum of Cost","","","128.74"]
		],
		"tabular_2row_1col_0data": [
				["","","Style","","",""],
				["Region","Gender","Fancy","Golf","Tee","Grand Total"],
				["East","Boy","","","",""],
				["","Girl","","","",""],
				["East Total","","","","",""],
				["West","Boy","","","",""],
				["","Girl","","","",""],
				["West Total","","","","",""],
				["Grand Total","","","","",""]
		],
		"tabular_2row_1col_1data": [
				["Sum of Price","","Style","","",""],
				["Region","Gender","Fancy","Golf","Tee","Grand Total"],
				["East","Boy","11.96","13","11.04","36"],
				["","Girl","13.74","12.12","11.27","37.13"],
				["East Total","","25.7","25.12","22.31","73.13"],
				["West","Boy","12.06","12.63","11.44","36.13"],
				["","Girl","","11.48","13.42","24.9"],
				["West Total","","12.06","24.11","24.86","61.03"],
				["Grand Total","","37.76","49.23","47.17","134.16"]
		],
		"tabular_2row_1col_2data_col": [
				["","","Style","Values","","","","","",""],
				["","","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
				["Region","Gender","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
				["East","Boy","11.96","11.74","13","12.6","11.04","10.42","36","34.76"],
				["","Girl","13.74","13.33","12.12","11.95","11.27","10.56","37.13","35.84"],
				["East Total","","25.7","25.07","25.12","24.55","22.31","20.98","73.13","70.6"],
				["West","Boy","12.06","11.51","12.63","11.73","11.44","10.94","36.13","34.18"],
				["","Girl","","","11.48","10.67","13.42","13.29","24.9","23.96"],
				["West Total","","12.06","11.51","24.11","22.4","24.86","24.23","61.03","58.14"],
				["Grand Total","","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"tabular_2row_1col_2data_row": [
				["","","","Style","","",""],
				["Region","Gender","Values","Fancy","Golf","Tee","Grand Total"],
				["East","Boy","Sum of Price","11.96","13","11.04","36"],
				["","","Sum of Cost","11.74","12.6","10.42","34.76"],
				["","Girl","Sum of Price","13.74","12.12","11.27","37.13"],
				["","","Sum of Cost","13.33","11.95","10.56","35.84"],
				["East Sum of Price","","","25.7","25.12","22.31","73.13"],
				["East Sum of Cost","","","25.07","24.55","20.98","70.6"],
				["West","Boy","Sum of Price","12.06","12.63","11.44","36.13"],
				["","","Sum of Cost","11.51","11.73","10.94","34.18"],
				["","Girl","Sum of Price","","11.48","13.42","24.9"],
				["","","Sum of Cost","","10.67","13.29","23.96"],
				["West Sum of Price","","","12.06","24.11","24.86","61.03"],
				["West Sum of Cost","","","11.51","22.4","24.23","58.14"],
				["Total Sum of Price","","","37.76","49.23","47.17","134.16"],
				["Total Sum of Cost","","","36.58","46.95","45.21","128.74"]
		],
		"tabular_2row_2col_0data": [
				["","","Style","Units","","","","","","","","","","","","",""],
				["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","Boy","","","","","","","","","","","","","","",""],
				["","Girl","","","","","","","","","","","","","","",""],
				["East Total","","","","","","","","","","","","","","","",""],
				["West","Boy","","","","","","","","","","","","","","",""],
				["","Girl","","","","","","","","","","","","","","",""],
				["West Total","","","","","","","","","","","","","","","",""],
				["Grand Total","","","","","","","","","","","","","","","",""]
		],
		"tabular_2row_2col_1data": [
				["Sum of Price","","Style","Units","","","","","","","","","","","","",""],
				["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["East Total","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["West","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["West Total","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["Grand Total","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"tabular_2row_2col_2data_col": [
				["","","Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["","","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
				["Region","Gender","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
				["East","Boy","","","","","11.96","11.74","11.96","11.74","","","","","13","12.6","","","13","12.6","","","","","11.04","10.42","","","11.04","10.42","36","34.76"],
				["","Girl","13.74","13.33","","","","","13.74","13.33","12.12","11.95","","","","","","","12.12","11.95","11.27","10.56","","","","","","","11.27","10.56","37.13","35.84"],
				["East Total","","13.74","13.33","","","11.96","11.74","25.7","25.07","12.12","11.95","","","13","12.6","","","25.12","24.55","11.27","10.56","","","11.04","10.42","","","22.31","20.98","73.13","70.6"],
				["West","Boy","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","","","12.63","11.73","","","11.44","10.94","","","","","11.44","10.94","36.13","34.18"],
				["","Girl","","","","","","","","","","","","","","","11.48","10.67","11.48","10.67","","","","","","","13.42","13.29","13.42","13.29","24.9","23.96"],
				["West Total","","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","11.48","10.67","24.11","22.4","","","11.44","10.94","","","13.42","13.29","24.86","24.23","61.03","58.14"],
				["Grand Total","","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"tabular_2row_2col_2data_col2": [
				["","","Style","Values","Units","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
				["","","Sum of Price","","","Sum of Cost","","","","","Sum of Price","","","","Sum of Cost","","","","","","Sum of Price","","","","Sum of Cost","","","","","","",""],
				["Region","Gender","10","11","12","10","11","12","","","10","11","12","15","10","11","12","15","","","10","11","12","15","10","11","12","15","","","",""],
				["East","Boy","","","11.96","","","11.74","11.96","11.74","","","13","","","","12.6","","13","12.6","","","11.04","","","","10.42","","11.04","10.42","36","34.76"],
				["","Girl","13.74","","","13.33","","","13.74","13.33","12.12","","","","11.95","","","","12.12","11.95","11.27","","","","10.56","","","","11.27","10.56","37.13","35.84"],
				["East Total","","13.74","","11.96","13.33","","11.74","25.7","25.07","12.12","","13","","11.95","","12.6","","25.12","24.55","11.27","","11.04","","10.56","","10.42","","22.31","20.98","73.13","70.6"],
				["West","Boy","","12.06","","","11.51","","12.06","11.51","","12.63","","","","11.73","","","12.63","11.73","","11.44","","","","10.94","","","11.44","10.94","36.13","34.18"],
				["","Girl","","","","","","","","","","","","11.48","","","","10.67","11.48","10.67","","","","13.42","","","","13.29","13.42","13.29","24.9","23.96"],
				["West Total","","","12.06","","","11.51","","12.06","11.51","","12.63","","11.48","","11.73","","10.67","24.11","22.4","","11.44","","13.42","","10.94","","13.29","24.86","24.23","61.03","58.14"],
				["Grand Total","","13.74","12.06","11.96","13.33","11.51","11.74","37.76","36.58","12.12","12.63","13","11.48","11.95","11.73","12.6","10.67","49.23","46.95","11.27","11.44","11.04","13.42","10.56","10.94","10.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"tabular_2row_2col_2data_col1": [
				["","","Values","Style","Units","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
				["","","Sum of Price","","","","","","","","","","","","","","Sum of Cost","","","","","","","","","","","","","","Total Sum of Price","Total Sum of Cost"],
				["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","",""],
				["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","","10","11","12","","10","11","12","15","","10","11","12","15","","",""],
				["East","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","36","34.76"],
				["","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","37.13","35.84"],
				["East Total","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","73.13","70.6"],
				["West","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","36.13","34.18"],
				["","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","","","","","","","","10.67","10.67","","","","13.29","13.29","24.9","23.96"],
				["West Total","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","61.03","58.14"],
				["Grand Total","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","134.16","128.74"]
		],
		"tabular_2row_2col_2data_row": [
				["","","","Style","Units","","","","","","","","","","","","",""],
				["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Gender","Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","Boy","Sum of Price","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["","","Sum of Cost","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["","Girl","Sum of Price","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["","","Sum of Cost","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["East Sum of Price","","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["East Sum of Cost","","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["West","Boy","Sum of Price","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["","","Sum of Cost","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["","Girl","Sum of Price","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["","","Sum of Cost","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["West Sum of Price","","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["West Sum of Cost","","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"tabular_2row_2col_2data_row2": [
				["","","","Style","Units","","","","","","","","","","","","",""],
				["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Region","Values","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","Sum of Price","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["","","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["","Sum of Cost","Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["","","Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["East Sum of Price","","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["East Sum of Cost","","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["West","Sum of Price","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["","","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["","Sum of Cost","Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["","","Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["West Sum of Price","","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["West Sum of Cost","","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"tabular_2row_2col_2data_row1": [
				["","","","Style","Units","","","","","","","","","","","","",""],
				["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["Values","Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["Sum of Price","East","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["","","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["","East Total","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["","West","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["","","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["","West Total","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["Sum of Cost","East","Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["","","Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["","East Total","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["","West","Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["","","Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["","West Total","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"gridDropZones_0row_0col_0data": [
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""],
			["","","","","","",""]
		],
		"gridDropZones_0row_0col_1data": [
			["Sum of Price","Total"],
			["Total","134.16"]
		],
		"gridDropZones_0row_0col_2data_col": [
			["","Values",""],
			["","Sum of Price","Sum of Cost"],
			["Total","134.16","128.74"]
		],
		"gridDropZones_0row_0col_2data_row": [
			["Values","Total"],
			["Sum of Price","134.16"],
			["Sum of Cost","128.74"]
		],
		"gridDropZones_0row_1col_0data": [
			["","Style","","",""],
			["","Fancy","Golf","Tee","Grand Total"],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""],
			["","","","",""]
		],
		"gridDropZones_0row_1col_1data": [
			["Sum of Price","Style","","",""],
			["","Fancy","Golf","Tee","Grand Total"],
			["Total","37.76","49.23","47.17","134.16"]
		],
		"gridDropZones_0row_1col_2data_col": [
			["","Style","Values","","","","","",""],
			["","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
			["","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
			["Total","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"gridDropZones_0row_1col_2data_row": [
			["","Style","","",""],
			["Values","Fancy","Golf","Tee","Grand Total"],
			["Sum of Price","37.76","49.23","47.17","134.16"],
			["Sum of Cost","36.58","46.95","45.21","128.74"]
		],
		"gridDropZones_0row_2col_0data": [
			["","Style","Units","","","","","","","","","","","","",""],
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","",""]
		],
		"gridDropZones_0row_2col_1data": [
			["Sum of Price","Style","Units","","","","","","","","","","","","",""],
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["Total","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"gridDropZones_0row_2col_2data_col": [
			["","Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
			["","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
			["","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
			["Total","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"gridDropZones_0row_2col_2data_row": [
			["","Style","Units","","","","","","","","","","","","",""],
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"gridDropZones_1row_0col_0data": [
			["","","","","","",""],
			["Region","","","","","",""],
			["East","","","","","",""],
			["West","","","","","",""],
			["Grand Total","","","","","",""]
		],
		"gridDropZones_1row_0col_1data": [
			["Sum of Price",""],
			["Region","Total"],
			["East","73.13"],
			["West","61.03"],
			["Grand Total","134.16"]
		],
		"gridDropZones_1row_0col_2data_col": [
			["","Values",""],
			["Region","Sum of Price","Sum of Cost"],
			["East","73.13","70.6"],
			["West","61.03","58.14"],
			["Grand Total","134.16","128.74"]
		],
		"gridDropZones_1row_0col_2data_row": [
			["Region","Values","Total"],
			["East","Sum of Price","73.13"],
			["","Sum of Cost","70.6"],
			["West","Sum of Price","61.03"],
			["","Sum of Cost","58.14"],
			["Total Sum of Price","","134.16"],
			["Total Sum of Cost","","128.74"]
		],
		"gridDropZones_1row_1col_0data": [
			["","Style","","",""],
			["Region","Fancy","Golf","Tee","Grand Total"],
			["East","","","",""],
			["West","","","",""],
			["Grand Total","","","",""]
		],
		"gridDropZones_1row_1col_1data": [
			["Sum of Price","Style","","",""],
			["Region","Fancy","Golf","Tee","Grand Total"],
			["East","25.7","25.12","22.31","73.13"],
			["West","12.06","24.11","24.86","61.03"],
			["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"gridDropZones_1row_1col_2data_col": [
			["","Style","Values","","","","","",""],
			["","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
			["Region","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
			["East","25.7","25.07","25.12","24.55","22.31","20.98","73.13","70.6"],
			["West","12.06","11.51","24.11","22.4","24.86","24.23","61.03","58.14"],
			["Grand Total","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"gridDropZones_1row_1col_2data_row": [
			["","","Style","","",""],
			["Region","Values","Fancy","Golf","Tee","Grand Total"],
			["East","Sum of Price","25.7","25.12","22.31","73.13"],
			["","Sum of Cost","25.07","24.55","20.98","70.6"],
			["West","Sum of Price","12.06","24.11","24.86","61.03"],
			["","Sum of Cost","11.51","22.4","24.23","58.14"],
			["Total Sum of Price","","37.76","49.23","47.17","134.16"],
			["Total Sum of Cost","","36.58","46.95","45.21","128.74"]
		],
		"gridDropZones_1row_2col_0data": [
			["","Style","Units","","","","","","","","","","","","",""],
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Region","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","","","","","","","","","","","","","","",""],
			["West","","","","","","","","","","","","","","",""],
			["Grand Total","","","","","","","","","","","","","","",""]
		],
		"gridDropZones_1row_2col_1data": [
			["Sum of Price","Style","Units","","","","","","","","","","","","",""],
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Region","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["Grand Total","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"gridDropZones_1row_2col_2data_col": [
			["","Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
			["","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
			["Region","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
			["East","13.74","13.33","","","11.96","11.74","25.7","25.07","12.12","11.95","","","13","12.6","","","25.12","24.55","11.27","10.56","","","11.04","10.42","","","22.31","20.98","73.13","70.6"],
			["West","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","11.48","10.67","24.11","22.4","","","11.44","10.94","","","13.42","13.29","24.86","24.23","61.03","58.14"],
			["Grand Total","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"gridDropZones_1row_2col_2data_row": [
			["","","Style","Units","","","","","","","","","","","","",""],
			["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Region","Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","Sum of Price","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["","Sum of Cost","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
			["West","Sum of Price","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["","Sum of Cost","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
			["Total Sum of Price","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Total Sum of Cost","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"gridDropZones_2row_0col_0data": [
			["","","","","","","",""],
			["Region","Gender","","","","","",""],
			["East","Boy","","","","","",""],
			["","Girl","","","","","",""],
			["East Total","","","","","","",""],
			["West","Boy","","","","","",""],
			["","Girl","","","","","",""],
			["West Total","","","","","","",""],
			["Grand Total","","","","","","",""]
		],
		"gridDropZones_2row_0col_1data": [
			["Sum of Price","",""],
			["Region","Gender","Total"],
			["East","Boy","36"],
			["","Girl","37.13"],
			["East Total","","73.13"],
			["West","Boy","36.13"],
			["","Girl","24.9"],
			["West Total","","61.03"],
			["Grand Total","","134.16"]
		],
		"gridDropZones_2row_0col_2data_col": [
			["","","Values",""],
			["Region","Gender","Sum of Price","Sum of Cost"],
			["East","Boy","36","34.76"],
			["","Girl","37.13","35.84"],
			["East Total","","73.13","70.6"],
			["West","Boy","36.13","34.18"],
			["","Girl","24.9","23.96"],
			["West Total","","61.03","58.14"],
			["Grand Total","","134.16","128.74"]
		],
		"gridDropZones_2row_0col_2data_row": [
			["Region","Gender","Values","Total"],
			["East","Boy","Sum of Price","36"],
			["","","Sum of Cost","34.76"],
			["","Girl","Sum of Price","37.13"],
			["","","Sum of Cost","35.84"],
			["East Sum of Price","","","73.13"],
			["East Sum of Cost","","","70.6"],
			["West","Boy","Sum of Price","36.13"],
			["","","Sum of Cost","34.18"],
			["","Girl","Sum of Price","24.9"],
			["","","Sum of Cost","23.96"],
			["West Sum of Price","","","61.03"],
			["West Sum of Cost","","","58.14"],
			["Total Sum of Price","","","134.16"],
			["Total Sum of Cost","","","128.74"]
		],
		"gridDropZones_2row_1col_0data": [
			["","","Style","","",""],
			["Region","Gender","Fancy","Golf","Tee","Grand Total"],
			["East","Boy","","","",""],
			["","Girl","","","",""],
			["East Total","","","","",""],
			["West","Boy","","","",""],
			["","Girl","","","",""],
			["West Total","","","","",""],
			["Grand Total","","","","",""]
		],
		"gridDropZones_2row_1col_1data": [
			["Sum of Price","","Style","","",""],
			["Region","Gender","Fancy","Golf","Tee","Grand Total"],
			["East","Boy","11.96","13","11.04","36"],
			["","Girl","13.74","12.12","11.27","37.13"],
			["East Total","","25.7","25.12","22.31","73.13"],
			["West","Boy","12.06","12.63","11.44","36.13"],
			["","Girl","","11.48","13.42","24.9"],
			["West Total","","12.06","24.11","24.86","61.03"],
			["Grand Total","","37.76","49.23","47.17","134.16"]
		],
		"gridDropZones_2row_1col_2data_col": [
			["","","Style","Values","","","","","",""],
			["","","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
			["Region","Gender","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
			["East","Boy","11.96","11.74","13","12.6","11.04","10.42","36","34.76"],
			["","Girl","13.74","13.33","12.12","11.95","11.27","10.56","37.13","35.84"],
			["East Total","","25.7","25.07","25.12","24.55","22.31","20.98","73.13","70.6"],
			["West","Boy","12.06","11.51","12.63","11.73","11.44","10.94","36.13","34.18"],
			["","Girl","","","11.48","10.67","13.42","13.29","24.9","23.96"],
			["West Total","","12.06","11.51","24.11","22.4","24.86","24.23","61.03","58.14"],
			["Grand Total","","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"gridDropZones_2row_1col_2data_row": [
			["","","","Style","","",""],
			["Region","Gender","Values","Fancy","Golf","Tee","Grand Total"],
			["East","Boy","Sum of Price","11.96","13","11.04","36"],
			["","","Sum of Cost","11.74","12.6","10.42","34.76"],
			["","Girl","Sum of Price","13.74","12.12","11.27","37.13"],
			["","","Sum of Cost","13.33","11.95","10.56","35.84"],
			["East Sum of Price","","","25.7","25.12","22.31","73.13"],
			["East Sum of Cost","","","25.07","24.55","20.98","70.6"],
			["West","Boy","Sum of Price","12.06","12.63","11.44","36.13"],
			["","","Sum of Cost","11.51","11.73","10.94","34.18"],
			["","Girl","Sum of Price","","11.48","13.42","24.9"],
			["","","Sum of Cost","","10.67","13.29","23.96"],
			["West Sum of Price","","","12.06","24.11","24.86","61.03"],
			["West Sum of Cost","","","11.51","22.4","24.23","58.14"],
			["Total Sum of Price","","","37.76","49.23","47.17","134.16"],
			["Total Sum of Cost","","","36.58","46.95","45.21","128.74"]
		],
		"gridDropZones_2row_2col_0data": [
			["","","Style","Units","","","","","","","","","","","","",""],
			["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","Boy","","","","","","","","","","","","","","",""],
			["","Girl","","","","","","","","","","","","","","",""],
			["East Total","","","","","","","","","","","","","","","",""],
			["West","Boy","","","","","","","","","","","","","","",""],
			["","Girl","","","","","","","","","","","","","","",""],
			["West Total","","","","","","","","","","","","","","","",""],
			["Grand Total","","","","","","","","","","","","","","","",""]
		],
		"gridDropZones_2row_2col_1data": [
			["Sum of Price","","Style","Units","","","","","","","","","","","","",""],
			["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
			["","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
			["East Total","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["West","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
			["","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
			["West Total","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["Grand Total","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"gridDropZones_2row_2col_2data_col": [
			["","","Style","Units","Values","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
			["","","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
			["Region","Gender","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
			["East","Boy","","","","","11.96","11.74","11.96","11.74","","","","","13","12.6","","","13","12.6","","","","","11.04","10.42","","","11.04","10.42","36","34.76"],
			["","Girl","13.74","13.33","","","","","13.74","13.33","12.12","11.95","","","","","","","12.12","11.95","11.27","10.56","","","","","","","11.27","10.56","37.13","35.84"],
			["East Total","","13.74","13.33","","","11.96","11.74","25.7","25.07","12.12","11.95","","","13","12.6","","","25.12","24.55","11.27","10.56","","","11.04","10.42","","","22.31","20.98","73.13","70.6"],
			["West","Boy","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","","","12.63","11.73","","","11.44","10.94","","","","","11.44","10.94","36.13","34.18"],
			["","Girl","","","","","","","","","","","","","","","11.48","10.67","11.48","10.67","","","","","","","13.42","13.29","13.42","13.29","24.9","23.96"],
			["West Total","","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","11.48","10.67","24.11","22.4","","","11.44","10.94","","","13.42","13.29","24.86","24.23","61.03","58.14"],
			["Grand Total","","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"gridDropZones_2row_2col_2data_col2": [
			["","","Style","Values","Units","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
			["","","Sum of Price","","","Sum of Cost","","","","","Sum of Price","","","","Sum of Cost","","","","","","Sum of Price","","","","Sum of Cost","","","","","","",""],
			["Region","Gender","10","11","12","10","11","12","","","10","11","12","15","10","11","12","15","","","10","11","12","15","10","11","12","15","","","",""],
			["East","Boy","","","11.96","","","11.74","11.96","11.74","","","13","","","","12.6","","13","12.6","","","11.04","","","","10.42","","11.04","10.42","36","34.76"],
			["","Girl","13.74","","","13.33","","","13.74","13.33","12.12","","","","11.95","","","","12.12","11.95","11.27","","","","10.56","","","","11.27","10.56","37.13","35.84"],
			["East Total","","13.74","","11.96","13.33","","11.74","25.7","25.07","12.12","","13","","11.95","","12.6","","25.12","24.55","11.27","","11.04","","10.56","","10.42","","22.31","20.98","73.13","70.6"],
			["West","Boy","","12.06","","","11.51","","12.06","11.51","","12.63","","","","11.73","","","12.63","11.73","","11.44","","","","10.94","","","11.44","10.94","36.13","34.18"],
			["","Girl","","","","","","","","","","","","11.48","","","","10.67","11.48","10.67","","","","13.42","","","","13.29","13.42","13.29","24.9","23.96"],
			["West Total","","","12.06","","","11.51","","12.06","11.51","","12.63","","11.48","","11.73","","10.67","24.11","22.4","","11.44","","13.42","","10.94","","13.29","24.86","24.23","61.03","58.14"],
			["Grand Total","","13.74","12.06","11.96","13.33","11.51","11.74","37.76","36.58","12.12","12.63","13","11.48","11.95","11.73","12.6","10.67","49.23","46.95","11.27","11.44","11.04","13.42","10.56","10.94","10.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"gridDropZones_2row_2col_2data_col1": [
			["","","Values","Style","Units","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","Sum of Price","","","","","","","","","","","","","","Sum of Cost","","","","","","","","","","","","","","Total Sum of Price","Total Sum of Cost"],
			["","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","",""],
			["Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","","10","11","12","","10","11","12","15","","10","11","12","15","","",""],
			["East","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","36","34.76"],
			["","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","37.13","35.84"],
			["East Total","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","73.13","70.6"],
			["West","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","36.13","34.18"],
			["","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","","","","","","","","10.67","10.67","","","","13.29","13.29","24.9","23.96"],
			["West Total","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","61.03","58.14"],
			["Grand Total","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","134.16","128.74"]
		],
		"gridDropZones_2row_2col_2data_row": [
			["","","","Style","Units","","","","","","","","","","","","",""],
			["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Region","Gender","Values","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","Boy","Sum of Price","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
			["","","Sum of Cost","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
			["","Girl","Sum of Price","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
			["","","Sum of Cost","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
			["East Sum of Price","","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["East Sum of Cost","","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
			["West","Boy","Sum of Price","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
			["","","Sum of Cost","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
			["","Girl","Sum of Price","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
			["","","Sum of Cost","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
			["West Sum of Price","","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["West Sum of Cost","","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
			["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"gridDropZones_2row_2col_2data_row2": [
			["","","","Style","Units","","","","","","","","","","","","",""],
			["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Region","Values","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","Sum of Price","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
			["","","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
			["","Sum of Cost","Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
			["","","Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
			["East Sum of Price","","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["East Sum of Cost","","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
			["West","Sum of Price","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
			["","","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
			["","Sum of Cost","Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
			["","","Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
			["West Sum of Price","","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["West Sum of Cost","","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
			["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"gridDropZones_2row_2col_2data_row1": [
			["","","","Style","Units","","","","","","","","","","","","",""],
			["","","","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["Values","Region","Gender","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["Sum of Price","East","Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
			["","","Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
			["","East Total","","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["","West","Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
			["","","Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
			["","West Total","","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["Sum of Cost","East","Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
			["","","Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
			["","East Total","","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
			["","West","Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
			["","","Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
			["","West Total","","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
			["Total Sum of Price","","","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Total Sum of Cost","","","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"showHeaders_0row_0col_0data": [
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""],
			["","",""]
		],
		"showHeaders_0row_0col_1data": [
			["Sum of Price"],
			["134.16"]
		],
		"showHeaders_0row_0col_2data_col": [
			["Sum of Price","Sum of Cost"],
			["134.16","128.74"]
		],
		"showHeaders_0row_0col_2data_row": [
			["Sum of Price","134.16"],
			["Sum of Cost","128.74"]
		],
		"showHeaders_0row_1col_0data": [
			["Fancy","Golf","Tee","Grand Total"]
		],
		"showHeaders_0row_1col_1data": [
			["","Fancy","Golf","Tee","Grand Total"],
			["Sum of Price","37.76","49.23","47.17","134.16"]
		],
		"showHeaders_0row_1col_2data_col": [
			["Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
			["Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
			["37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"showHeaders_0row_1col_2data_row": [
			["","Fancy","Golf","Tee","Grand Total"],
			["Sum of Price","37.76","49.23","47.17","134.16"],
			["Sum of Cost","36.58","46.95","45.21","128.74"]
		],
		"showHeaders_0row_2col_0data": [
			["Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["10","11","12","","10","11","12","15","","10","11","12","15","",""]
		],
		"showHeaders_0row_2col_1data": [
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"showHeaders_0row_2col_2data_col": [
			["Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
			["10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
			["Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
			["13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"showHeaders_0row_2col_2data_row": [
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"showHeaders_1row_0col_0data": [
			["East"],
			["West"],
			["Grand Total"]
		],
		"showHeaders_1row_0col_1data": [
			["","Sum of Price"],
			["East","73.13"],
			["West","61.03"],
			["Grand Total","134.16"]
		],
		"showHeaders_1row_0col_2data_col": [
			["","Sum of Price","Sum of Cost"],
			["East","73.13","70.6"],
			["West","61.03","58.14"],
			["Grand Total","134.16","128.74"]
		],
		"showHeaders_1row_0col_2data_row": [
			["East",""],
			["Sum of Price","73.13"],
			["Sum of Cost","70.6"],
			["West",""],
			["Sum of Price","61.03"],
			["Sum of Cost","58.14"],
			["Total Sum of Price","134.16"],
			["Total Sum of Cost","128.74"]
		],
		"showHeaders_1row_1col_0data": [
			["","Fancy","Golf","Tee","Grand Total"],
			["East","","","",""],
			["West","","","",""],
			["Grand Total","","","",""]
		],
		"showHeaders_1row_1col_1data": [
			["Sum of Price","","","",""],
			["","Fancy","Golf","Tee","Grand Total"],
			["East","25.7","25.12","22.31","73.13"],
			["West","12.06","24.11","24.86","61.03"],
			["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"showHeaders_1row_1col_2data_col": [
			["","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
			["","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
			["East","25.7","25.07","25.12","24.55","22.31","20.98","73.13","70.6"],
			["West","12.06","11.51","24.11","22.4","24.86","24.23","61.03","58.14"],
			["Grand Total","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"showHeaders_1row_1col_2data_row": [
			["","Fancy","Golf","Tee","Grand Total"],
			["East","","","",""],
			["Sum of Price","25.7","25.12","22.31","73.13"],
			["Sum of Cost","25.07","24.55","20.98","70.6"],
			["West","","","",""],
			["Sum of Price","12.06","24.11","24.86","61.03"],
			["Sum of Cost","11.51","22.4","24.23","58.14"],
			["Total Sum of Price","37.76","49.23","47.17","134.16"],
			["Total Sum of Cost","36.58","46.95","45.21","128.74"]
		],
		"showHeaders_1row_2col_0data": [
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","","","","","","","","","","","","","","",""],
			["West","","","","","","","","","","","","","","",""],
			["Grand Total","","","","","","","","","","","","","","",""]
		],
		"showHeaders_1row_2col_1data": [
			["Sum of Price","","","","","","","","","","","","","","",""],
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["Grand Total","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"showHeaders_1row_2col_2data_col": [
			["","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
			["","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
			["","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
			["East","13.74","13.33","","","11.96","11.74","25.7","25.07","12.12","11.95","","","13","12.6","","","25.12","24.55","11.27","10.56","","","11.04","10.42","","","22.31","20.98","73.13","70.6"],
			["West","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","11.48","10.67","24.11","22.4","","","11.44","10.94","","","13.42","13.29","24.86","24.23","61.03","58.14"],
			["Grand Total","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"showHeaders_1row_2col_2data_row": [
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","","","","","","","","","","","","","","",""],
			["Sum of Price","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["Sum of Cost","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
			["West","","","","","","","","","","","","","","",""],
			["Sum of Price","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["Sum of Cost","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
			["Total Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Total Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"showHeaders_2row_0col_0data": [
			["East"],
			["Boy"],
			["Girl"],
			["West"],
			["Boy"],
			["Girl"],
			["Grand Total"]
		],
		"showHeaders_2row_0col_1data": [
			["","Sum of Price"],
			["East","73.13"],
			["Boy","36"],
			["Girl","37.13"],
			["West","61.03"],
			["Boy","36.13"],
			["Girl","24.9"],
			["Grand Total","134.16"]
		],
		"showHeaders_2row_0col_2data_col": [
			["","Sum of Price","Sum of Cost"],
			["East","73.13","70.6"],
			["Boy","36","34.76"],
			["Girl","37.13","35.84"],
			["West","61.03","58.14"],
			["Boy","36.13","34.18"],
			["Girl","24.9","23.96"],
			["Grand Total","134.16","128.74"]
		],
		"showHeaders_2row_0col_2data_row": [
			["East",""],
			["Boy",""],
			["Sum of Price","36"],
			["Sum of Cost","34.76"],
			["Girl",""],
			["Sum of Price","37.13"],
			["Sum of Cost","35.84"],
			["East Sum of Price","73.13"],
			["East Sum of Cost","70.6"],
			["West",""],
			["Boy",""],
			["Sum of Price","36.13"],
			["Sum of Cost","34.18"],
			["Girl",""],
			["Sum of Price","24.9"],
			["Sum of Cost","23.96"],
			["West Sum of Price","61.03"],
			["West Sum of Cost","58.14"],
			["Total Sum of Price","134.16"],
			["Total Sum of Cost","128.74"]
		],
		"showHeaders_2row_1col_0data": [
			["","Fancy","Golf","Tee","Grand Total"],
			["East","","","",""],
			["Boy","","","",""],
			["Girl","","","",""],
			["West","","","",""],
			["Boy","","","",""],
			["Girl","","","",""],
			["Grand Total","","","",""]
		],
		"showHeaders_2row_1col_1data": [
			["Sum of Price","","","",""],
			["","Fancy","Golf","Tee","Grand Total"],
			["East","25.7","25.12","22.31","73.13"],
			["Boy","11.96","13","11.04","36"],
			["Girl","13.74","12.12","11.27","37.13"],
			["West","12.06","24.11","24.86","61.03"],
			["Boy","12.06","12.63","11.44","36.13"],
			["Girl","","11.48","13.42","24.9"],
			["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"showHeaders_2row_1col_2data_col": [
			["","Fancy","","Golf","","Tee","","Total Sum of Price","Total Sum of Cost"],
			["","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
			["East","25.7","25.07","25.12","24.55","22.31","20.98","73.13","70.6"],
			["Boy","11.96","11.74","13","12.6","11.04","10.42","36","34.76"],
			["Girl","13.74","13.33","12.12","11.95","11.27","10.56","37.13","35.84"],
			["West","12.06","11.51","24.11","22.4","24.86","24.23","61.03","58.14"],
			["Boy","12.06","11.51","12.63","11.73","11.44","10.94","36.13","34.18"],
			["Girl","","","11.48","10.67","13.42","13.29","24.9","23.96"],
			["Grand Total","37.76","36.58","49.23","46.95","47.17","45.21","134.16","128.74"]
		],
		"showHeaders_2row_1col_2data_row": [
			["","Fancy","Golf","Tee","Grand Total"],
			["East","","","",""],
			["Boy","","","",""],
			["Sum of Price","11.96","13","11.04","36"],
			["Sum of Cost","11.74","12.6","10.42","34.76"],
			["Girl","","","",""],
			["Sum of Price","13.74","12.12","11.27","37.13"],
			["Sum of Cost","13.33","11.95","10.56","35.84"],
			["East Sum of Price","25.7","25.12","22.31","73.13"],
			["East Sum of Cost","25.07","24.55","20.98","70.6"],
			["West","","","",""],
			["Boy","","","",""],
			["Sum of Price","12.06","12.63","11.44","36.13"],
			["Sum of Cost","11.51","11.73","10.94","34.18"],
			["Girl","","","",""],
			["Sum of Price","","11.48","13.42","24.9"],
			["Sum of Cost","","10.67","13.29","23.96"],
			["West Sum of Price","12.06","24.11","24.86","61.03"],
			["West Sum of Cost","11.51","22.4","24.23","58.14"],
			["Total Sum of Price","37.76","49.23","47.17","134.16"],
			["Total Sum of Cost","36.58","46.95","45.21","128.74"]
		],
		"showHeaders_2row_2col_0data": [
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","","","","","","","","","","","","","","",""],
			["Boy","","","","","","","","","","","","","","",""],
			["Girl","","","","","","","","","","","","","","",""],
			["West","","","","","","","","","","","","","","",""],
			["Boy","","","","","","","","","","","","","","",""],
			["Girl","","","","","","","","","","","","","","",""],
			["Grand Total","","","","","","","","","","","","","","",""]
		],
		"showHeaders_2row_2col_1data": [
			["Sum of Price","","","","","","","","","","","","","","",""],
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
			["Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
			["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
			["Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
			["Grand Total","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"]
		],
		"showHeaders_2row_2col_2data_col": [
			["","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
			["","10","","11","","12","","","","10","","11","","12","","15","","","","10","","11","","12","","15","","","","",""],
			["","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","","","",""],
			["East","13.74","13.33","","","11.96","11.74","25.7","25.07","12.12","11.95","","","13","12.6","","","25.12","24.55","11.27","10.56","","","11.04","10.42","","","22.31","20.98","73.13","70.6"],
			["Boy","","","","","11.96","11.74","11.96","11.74","","","","","13","12.6","","","13","12.6","","","","","11.04","10.42","","","11.04","10.42","36","34.76"],
			["Girl","13.74","13.33","","","","","13.74","13.33","12.12","11.95","","","","","","","12.12","11.95","11.27","10.56","","","","","","","11.27","10.56","37.13","35.84"],
			["West","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","11.48","10.67","24.11","22.4","","","11.44","10.94","","","13.42","13.29","24.86","24.23","61.03","58.14"],
			["Boy","","","12.06","11.51","","","12.06","11.51","","","12.63","11.73","","","","","12.63","11.73","","","11.44","10.94","","","","","11.44","10.94","36.13","34.18"],
			["Girl","","","","","","","","","","","","","","","11.48","10.67","11.48","10.67","","","","","","","13.42","13.29","13.42","13.29","24.9","23.96"],
			["Grand Total","13.74","13.33","12.06","11.51","11.96","11.74","37.76","36.58","12.12","11.95","12.63","11.73","13","12.6","11.48","10.67","49.23","46.95","11.27","10.56","11.44","10.94","11.04","10.42","13.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"showHeaders_2row_2col_2data_col2": [
			["","Column Labels","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","Fancy","","","","","","Fancy Sum of Price","Fancy Sum of Cost","Golf","","","","","","","","Golf Sum of Price","Golf Sum of Cost","Tee","","","","","","","","Tee Sum of Price","Tee Sum of Cost","Total Sum of Price","Total Sum of Cost"],
			["","Sum of Price","","","Sum of Cost","","","","","Sum of Price","","","","Sum of Cost","","","","","","Sum of Price","","","","Sum of Cost","","","","","","",""],
			["Row Labels","10","11","12","10","11","12","","","10","11","12","15","10","11","12","15","","","10","11","12","15","10","11","12","15","","","",""],
			["East","13.74","","11.96","13.33","","11.74","25.7","25.07","12.12","","13","","11.95","","12.6","","25.12","24.55","11.27","","11.04","","10.56","","10.42","","22.31","20.98","73.13","70.6"],
			["Boy","","","11.96","","","11.74","11.96","11.74","","","13","","","","12.6","","13","12.6","","","11.04","","","","10.42","","11.04","10.42","36","34.76"],
			["Girl","13.74","","","13.33","","","13.74","13.33","12.12","","","","11.95","","","","12.12","11.95","11.27","","","","10.56","","","","11.27","10.56","37.13","35.84"],
			["West","","12.06","","","11.51","","12.06","11.51","","12.63","","11.48","","11.73","","10.67","24.11","22.4","","11.44","","13.42","","10.94","","13.29","24.86","24.23","61.03","58.14"],
			["Boy","","12.06","","","11.51","","12.06","11.51","","12.63","","","","11.73","","","12.63","11.73","","11.44","","","","10.94","","","11.44","10.94","36.13","34.18"],
			["Girl","","","","","","","","","","","","11.48","","","","10.67","11.48","10.67","","","","13.42","","","","13.29","13.42","13.29","24.9","23.96"],
			["Grand Total","13.74","12.06","11.96","13.33","11.51","11.74","37.76","36.58","12.12","12.63","13","11.48","11.95","11.73","12.6","10.67","49.23","46.95","11.27","11.44","11.04","13.42","10.56","10.94","10.42","13.29","47.17","45.21","134.16","128.74"]
		],
		"showHeaders_2row_2col_2data_col1": [
			["","Sum of Price","","","","","","","","","","","","","","Sum of Cost","","","","","","","","","","","","","","Total Sum of Price","Total Sum of Cost"],
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","",""],
			["","10","11","12","","10","11","12","15","","10","11","12","15","","10","11","12","","10","11","12","15","","10","11","12","15","","",""],
			["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","73.13","70.6"],
			["Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","36","34.76"],
			["Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","37.13","35.84"],
			["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","61.03","58.14"],
			["Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","36.13","34.18"],
			["Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","","","","","","","","10.67","10.67","","","","13.29","13.29","24.9","23.96"],
			["Grand Total","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","134.16","128.74"]
		],
		"showHeaders_2row_2col_2data_row": [
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
				["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
				["East","","","","","","","","","","","","","","",""],
				["Boy","","","","","","","","","","","","","","",""],
				["Sum of Price","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
				["Sum of Cost","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
				["Girl","","","","","","","","","","","","","","",""],
				["Sum of Price","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
				["Sum of Cost","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
				["East Sum of Price","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
				["East Sum of Cost","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
				["West","","","","","","","","","","","","","","",""],
				["Boy","","","","","","","","","","","","","","",""],
				["Sum of Price","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
				["Sum of Cost","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
				["Girl","","","","","","","","","","","","","","",""],
				["Sum of Price","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
				["Sum of Cost","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
				["West Sum of Price","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
				["West Sum of Cost","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
				["Total Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
				["Total Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
			],
		"showHeaders_2row_2col_2data_row2": [
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["East","","","","","","","","","","","","","","",""],
			["Sum of Price","","","","","","","","","","","","","","",""],
			["Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
			["Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
			["Sum of Cost","","","","","","","","","","","","","","",""],
			["Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
			["Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
			["East Sum of Price","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["East Sum of Cost","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
			["West","","","","","","","","","","","","","","",""],
			["Sum of Price","","","","","","","","","","","","","","",""],
			["Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
			["Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
			["Sum of Cost","","","","","","","","","","","","","","",""],
			["Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
			["Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
			["West Sum of Price","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["West Sum of Cost","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
			["Total Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Total Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"showHeaders_2row_2col_2data_row1": [
			["","Fancy","","","Fancy Total","Golf","","","","Golf Total","Tee","","","","Tee Total","Grand Total"],
			["","10","11","12","","10","11","12","15","","10","11","12","15","",""],
			["Sum of Price","","","","","","","","","","","","","","",""],
			["East","13.74","","11.96","25.7","12.12","","13","","25.12","11.27","","11.04","","22.31","73.13"],
			["Boy","","","11.96","11.96","","","13","","13","","","11.04","","11.04","36"],
			["Girl","13.74","","","13.74","12.12","","","","12.12","11.27","","","","11.27","37.13"],
			["West","","12.06","","12.06","","12.63","","11.48","24.11","","11.44","","13.42","24.86","61.03"],
			["Boy","","12.06","","12.06","","12.63","","","12.63","","11.44","","","11.44","36.13"],
			["Girl","","","","","","","","11.48","11.48","","","","13.42","13.42","24.9"],
			["Sum of Cost","","","","","","","","","","","","","","",""],
			["East","13.33","","11.74","25.07","11.95","","12.6","","24.55","10.56","","10.42","","20.98","70.6"],
			["Boy","","","11.74","11.74","","","12.6","","12.6","","","10.42","","10.42","34.76"],
			["Girl","13.33","","","13.33","11.95","","","","11.95","10.56","","","","10.56","35.84"],
			["West","","11.51","","11.51","","11.73","","10.67","22.4","","10.94","","13.29","24.23","58.14"],
			["Boy","","11.51","","11.51","","11.73","","","11.73","","10.94","","","10.94","34.18"],
			["Girl","","","","","","","","10.67","10.67","","","","13.29","13.29","23.96"],
			["Total Sum of Price","13.74","12.06","11.96","37.76","12.12","12.63","13","11.48","49.23","11.27","11.44","11.04","13.42","47.17","134.16"],
			["Total Sum of Cost","13.33","11.51","11.74","36.58","11.95","11.73","12.6","10.67","46.95","10.56","10.94","10.42","13.29","45.21","128.74"]
		],
		"subtotal_compact_none": [
				["Row Labels",""],
				["East",""],
				["Boy",""],
				["Sum of Price",""],
				["Fancy",""],
				["12","11.96"],
				["Golf",""],
				["12","13"],
				["Tee",""],
				["12","11.04"],
				["Sum of Cost",""],
				["Fancy",""],
				["12","11.74"],
				["Golf",""],
				["12","12.6"],
				["Tee",""],
				["12","10.42"],
				["Girl",""],
				["Sum of Price",""],
				["Fancy",""],
				["10","13.74"],
				["Golf",""],
				["10","12.12"],
				["Tee",""],
				["10","11.27"],
				["Sum of Cost",""],
				["Fancy",""],
				["10","13.33"],
				["Golf",""],
				["10","11.95"],
				["Tee",""],
				["10","10.56"],
				["West",""],
				["Boy",""],
				["Sum of Price",""],
				["Fancy",""],
				["11","12.06"],
				["Golf",""],
				["11","12.63"],
				["Tee",""],
				["11","11.44"],
				["Sum of Cost",""],
				["Fancy",""],
				["11","11.51"],
				["Golf",""],
				["11","11.73"],
				["Tee",""],
				["11","10.94"],
				["Girl",""],
				["Sum of Price",""],
				["Golf",""],
				["15","11.48"],
				["Tee",""],
				["15","13.42"],
				["Sum of Cost",""],
				["Golf",""],
				["15","10.67"],
				["Tee",""],
				["15","13.29"],
				["Total Sum of Price","134.16"],
				["Total Sum of Cost","128.74"]
		],
		"subtotal_compact_top": [
				["Row Labels",""],
				["East",""],
				["Boy",""],
				["Sum of Price",""],
				["Fancy","11.96"],
				["12","11.96"],
				["Golf","13"],
				["12","13"],
				["Tee","11.04"],
				["12","11.04"],
				["Sum of Cost",""],
				["Fancy","11.74"],
				["12","11.74"],
				["Golf","12.6"],
				["12","12.6"],
				["Tee","10.42"],
				["12","10.42"],
				["Boy Sum of Price","36"],
				["Boy Sum of Cost","34.76"],
				["Girl",""],
				["Sum of Price",""],
				["Fancy","13.74"],
				["10","13.74"],
				["Golf","12.12"],
				["10","12.12"],
				["Tee","11.27"],
				["10","11.27"],
				["Sum of Cost",""],
				["Fancy","13.33"],
				["10","13.33"],
				["Golf","11.95"],
				["10","11.95"],
				["Tee","10.56"],
				["10","10.56"],
				["Girl Sum of Price","37.13"],
				["Girl Sum of Cost","35.84"],
				["East Sum of Price","73.13"],
				["East Sum of Cost","70.6"],
				["West",""],
				["Boy",""],
				["Sum of Price",""],
				["Fancy","12.06"],
				["11","12.06"],
				["Golf","12.63"],
				["11","12.63"],
				["Tee","11.44"],
				["11","11.44"],
				["Sum of Cost",""],
				["Fancy","11.51"],
				["11","11.51"],
				["Golf","11.73"],
				["11","11.73"],
				["Tee","10.94"],
				["11","10.94"],
				["Boy Sum of Price","36.13"],
				["Boy Sum of Cost","34.18"],
				["Girl",""],
				["Sum of Price",""],
				["Golf","11.48"],
				["15","11.48"],
				["Tee","13.42"],
				["15","13.42"],
				["Sum of Cost",""],
				["Golf","10.67"],
				["15","10.67"],
				["Tee","13.29"],
				["15","13.29"],
				["Girl Sum of Price","24.9"],
				["Girl Sum of Cost","23.96"],
				["West Sum of Price","61.03"],
				["West Sum of Cost","58.14"],
				["Total Sum of Price","134.16"],
				["Total Sum of Cost","128.74"]
		],
		"subtotal_compact_bottom": [
				["Row Labels",""],
				["East",""],
				["Boy",""],
				["Sum of Price",""],
				["Fancy",""],
				["12","11.96"],
				["Fancy Total","11.96"],
				["Golf",""],
				["12","13"],
				["Golf Total","13"],
				["Tee",""],
				["12","11.04"],
				["Tee Total","11.04"],
				["Sum of Cost",""],
				["Fancy",""],
				["12","11.74"],
				["Fancy Total","11.74"],
				["Golf",""],
				["12","12.6"],
				["Golf Total","12.6"],
				["Tee",""],
				["12","10.42"],
				["Tee Total","10.42"],
				["Boy Sum of Price","36"],
				["Boy Sum of Cost","34.76"],
				["Girl",""],
				["Sum of Price",""],
				["Fancy",""],
				["10","13.74"],
				["Fancy Total","13.74"],
				["Golf",""],
				["10","12.12"],
				["Golf Total","12.12"],
				["Tee",""],
				["10","11.27"],
				["Tee Total","11.27"],
				["Sum of Cost",""],
				["Fancy",""],
				["10","13.33"],
				["Fancy Total","13.33"],
				["Golf",""],
				["10","11.95"],
				["Golf Total","11.95"],
				["Tee",""],
				["10","10.56"],
				["Tee Total","10.56"],
				["Girl Sum of Price","37.13"],
				["Girl Sum of Cost","35.84"],
				["East Sum of Price","73.13"],
				["East Sum of Cost","70.6"],
				["West",""],
				["Boy",""],
				["Sum of Price",""],
				["Fancy",""],
				["11","12.06"],
				["Fancy Total","12.06"],
				["Golf",""],
				["11","12.63"],
				["Golf Total","12.63"],
				["Tee",""],
				["11","11.44"],
				["Tee Total","11.44"],
				["Sum of Cost",""],
				["Fancy",""],
				["11","11.51"],
				["Fancy Total","11.51"],
				["Golf",""],
				["11","11.73"],
				["Golf Total","11.73"],
				["Tee",""],
				["11","10.94"],
				["Tee Total","10.94"],
				["Boy Sum of Price","36.13"],
				["Boy Sum of Cost","34.18"],
				["Girl",""],
				["Sum of Price",""],
				["Golf",""],
				["15","11.48"],
				["Golf Total","11.48"],
				["Tee",""],
				["15","13.42"],
				["Tee Total","13.42"],
				["Sum of Cost",""],
				["Golf",""],
				["15","10.67"],
				["Golf Total","10.67"],
				["Tee",""],
				["15","13.29"],
				["Tee Total","13.29"],
				["Girl Sum of Price","24.9"],
				["Girl Sum of Cost","23.96"],
				["West Sum of Price","61.03"],
				["West Sum of Cost","58.14"],
				["Total Sum of Price","134.16"],
				["Total Sum of Cost","128.74"]
		],
		"subtotal_tabular_none": [
				["Region","Gender","Values","Style","Units",""],
				["East","Boy","Sum of Price","Fancy","12","11.96"],
				["","","","Golf","12","13"],
				["","","","Tee","12","11.04"],
				["","","Sum of Cost","Fancy","12","11.74"],
				["","","","Golf","12","12.6"],
				["","","","Tee","12","10.42"],
				["","Girl","Sum of Price","Fancy","10","13.74"],
				["","","","Golf","10","12.12"],
				["","","","Tee","10","11.27"],
				["","","Sum of Cost","Fancy","10","13.33"],
				["","","","Golf","10","11.95"],
				["","","","Tee","10","10.56"],
				["West","Boy","Sum of Price","Fancy","11","12.06"],
				["","","","Golf","11","12.63"],
				["","","","Tee","11","11.44"],
				["","","Sum of Cost","Fancy","11","11.51"],
				["","","","Golf","11","11.73"],
				["","","","Tee","11","10.94"],
				["","Girl","Sum of Price","Golf","15","11.48"],
				["","","","Tee","15","13.42"],
				["","","Sum of Cost","Golf","15","10.67"],
				["","","","Tee","15","13.29"],
				["Total Sum of Price","","","","","134.16"],
				["Total Sum of Cost","","","","","128.74"]
		],
		"subtotal_tabular_top": [
				["Region","Gender","Values","Style","Units",""],
				["East","Boy","Sum of Price","Fancy","12","11.96"],
				["","","","Fancy Total","","11.96"],
				["","","","Golf","12","13"],
				["","","","Golf Total","","13"],
				["","","","Tee","12","11.04"],
				["","","","Tee Total","","11.04"],
				["","","Sum of Cost","Fancy","12","11.74"],
				["","","","Fancy Total","","11.74"],
				["","","","Golf","12","12.6"],
				["","","","Golf Total","","12.6"],
				["","","","Tee","12","10.42"],
				["","","","Tee Total","","10.42"],
				["","Boy Sum of Price","","","","36"],
				["","Boy Sum of Cost","","","","34.76"],
				["","Girl","Sum of Price","Fancy","10","13.74"],
				["","","","Fancy Total","","13.74"],
				["","","","Golf","10","12.12"],
				["","","","Golf Total","","12.12"],
				["","","","Tee","10","11.27"],
				["","","","Tee Total","","11.27"],
				["","","Sum of Cost","Fancy","10","13.33"],
				["","","","Fancy Total","","13.33"],
				["","","","Golf","10","11.95"],
				["","","","Golf Total","","11.95"],
				["","","","Tee","10","10.56"],
				["","","","Tee Total","","10.56"],
				["","Girl Sum of Price","","","","37.13"],
				["","Girl Sum of Cost","","","","35.84"],
				["East Sum of Price","","","","","73.13"],
				["East Sum of Cost","","","","","70.6"],
				["West","Boy","Sum of Price","Fancy","11","12.06"],
				["","","","Fancy Total","","12.06"],
				["","","","Golf","11","12.63"],
				["","","","Golf Total","","12.63"],
				["","","","Tee","11","11.44"],
				["","","","Tee Total","","11.44"],
				["","","Sum of Cost","Fancy","11","11.51"],
				["","","","Fancy Total","","11.51"],
				["","","","Golf","11","11.73"],
				["","","","Golf Total","","11.73"],
				["","","","Tee","11","10.94"],
				["","","","Tee Total","","10.94"],
				["","Boy Sum of Price","","","","36.13"],
				["","Boy Sum of Cost","","","","34.18"],
				["","Girl","Sum of Price","Golf","15","11.48"],
				["","","","Golf Total","","11.48"],
				["","","","Tee","15","13.42"],
				["","","","Tee Total","","13.42"],
				["","","Sum of Cost","Golf","15","10.67"],
				["","","","Golf Total","","10.67"],
				["","","","Tee","15","13.29"],
				["","","","Tee Total","","13.29"],
				["","Girl Sum of Price","","","","24.9"],
				["","Girl Sum of Cost","","","","23.96"],
				["West Sum of Price","","","","","61.03"],
				["West Sum of Cost","","","","","58.14"],
				["Total Sum of Price","","","","","134.16"],
				["Total Sum of Cost","","","","","128.74"]
		],
		"subtotal_tabular_bottom": [
				["Region","Gender","Values","Style","Units",""],
				["East","Boy","Sum of Price","Fancy","12","11.96"],
				["","","","Fancy Total","","11.96"],
				["","","","Golf","12","13"],
				["","","","Golf Total","","13"],
				["","","","Tee","12","11.04"],
				["","","","Tee Total","","11.04"],
				["","","Sum of Cost","Fancy","12","11.74"],
				["","","","Fancy Total","","11.74"],
				["","","","Golf","12","12.6"],
				["","","","Golf Total","","12.6"],
				["","","","Tee","12","10.42"],
				["","","","Tee Total","","10.42"],
				["","Boy Sum of Price","","","","36"],
				["","Boy Sum of Cost","","","","34.76"],
				["","Girl","Sum of Price","Fancy","10","13.74"],
				["","","","Fancy Total","","13.74"],
				["","","","Golf","10","12.12"],
				["","","","Golf Total","","12.12"],
				["","","","Tee","10","11.27"],
				["","","","Tee Total","","11.27"],
				["","","Sum of Cost","Fancy","10","13.33"],
				["","","","Fancy Total","","13.33"],
				["","","","Golf","10","11.95"],
				["","","","Golf Total","","11.95"],
				["","","","Tee","10","10.56"],
				["","","","Tee Total","","10.56"],
				["","Girl Sum of Price","","","","37.13"],
				["","Girl Sum of Cost","","","","35.84"],
				["East Sum of Price","","","","","73.13"],
				["East Sum of Cost","","","","","70.6"],
				["West","Boy","Sum of Price","Fancy","11","12.06"],
				["","","","Fancy Total","","12.06"],
				["","","","Golf","11","12.63"],
				["","","","Golf Total","","12.63"],
				["","","","Tee","11","11.44"],
				["","","","Tee Total","","11.44"],
				["","","Sum of Cost","Fancy","11","11.51"],
				["","","","Fancy Total","","11.51"],
				["","","","Golf","11","11.73"],
				["","","","Golf Total","","11.73"],
				["","","","Tee","11","10.94"],
				["","","","Tee Total","","10.94"],
				["","Boy Sum of Price","","","","36.13"],
				["","Boy Sum of Cost","","","","34.18"],
				["","Girl","Sum of Price","Golf","15","11.48"],
				["","","","Golf Total","","11.48"],
				["","","","Tee","15","13.42"],
				["","","","Tee Total","","13.42"],
				["","","Sum of Cost","Golf","15","10.67"],
				["","","","Golf Total","","10.67"],
				["","","","Tee","15","13.29"],
				["","","","Tee Total","","13.29"],
				["","Girl Sum of Price","","","","24.9"],
				["","Girl Sum of Cost","","","","23.96"],
				["West Sum of Price","","","","","61.03"],
				["West Sum of Cost","","","","","58.14"],
				["Total Sum of Price","","","","","134.16"],
				["Total Sum of Cost","","","","","128.74"]
		],
		"insertBlankRow_1row":[
				["Row Labels","Sum of Price","Sum of Cost"],
				["East","73.13","70.6"],
				["West","61.03","58.14"],
				["Grand Total","134.16","128.74"]
		],
		"insertBlankRow_2row":[
				["Row Labels",""],
				["East",""],
				["Sum of Price","73.13"],
				["Sum of Cost","70.6"],
				["",""],
				["West",""],
				["Sum of Price","61.03"],
				["Sum of Cost","58.14"],
				["",""],
				["Total Sum of Price","134.16"],
				["Total Sum of Cost","128.74"]
		],
		"insertBlankRow_3row":[
				["Row Labels",""],
				["East",""],
				["Sum of Price",""],
				["Boy","36"],
				["Girl","37.13"],
				["Sum of Cost",""],
				["Boy","34.76"],
				["Girl","35.84"],
				["East Sum of Price","73.13"],
				["East Sum of Cost","70.6"],
				["",""],
				["West",""],
				["Sum of Price",""],
				["Boy","36.13"],
				["Girl","24.9"],
				["Sum of Cost",""],
				["Boy","34.18"],
				["Girl","23.96"],
				["West Sum of Price","61.03"],
				["West Sum of Cost","58.14"],
				["",""],
				["Total Sum of Price","134.16"],
				["Total Sum of Cost","128.74"]
		],
		"insertBlankRow_4row":[
				["Row Labels",""],
				["East",""],
				["Sum of Price",""],
				["Boy","36"],
				["Fancy","11.96"],
				["Golf","13"],
				["Tee","11.04"],
				["",""],
				["Girl","37.13"],
				["Fancy","13.74"],
				["Golf","12.12"],
				["Tee","11.27"],
				["",""],
				["Sum of Cost",""],
				["Boy","34.76"],
				["Fancy","11.74"],
				["Golf","12.6"],
				["Tee","10.42"],
				["",""],
				["Girl","35.84"],
				["Fancy","13.33"],
				["Golf","11.95"],
				["Tee","10.56"],
				["",""],
				["East Sum of Price","73.13"],
				["East Sum of Cost","70.6"],
				["",""],
				["West",""],
				["Sum of Price",""],
				["Boy","36.13"],
				["Fancy","12.06"],
				["Golf","12.63"],
				["Tee","11.44"],
				["",""],
				["Girl","24.9"],
				["Golf","11.48"],
				["Tee","13.42"],
				["",""],
				["Sum of Cost",""],
				["Boy","34.18"],
				["Fancy","11.51"],
				["Golf","11.73"],
				["Tee","10.94"],
				["",""],
				["Girl","23.96"],
				["Golf","10.67"],
				["Tee","13.29"],
				["",""],
				["West Sum of Price","61.03"],
				["West Sum of Cost","58.14"],
				["",""],
				["Total Sum of Price","134.16"],
				["Total Sum of Cost","128.74"]
		],
		"filter_downThenOver1":[
			["Region","(All)"],
		],
		"filter_downThenOver3":[
			["Region","(All)"],
			["Gender","(All)"],
			["Style","(All)"],
		],
		"filter_downThenOver3_2wrap":[
			["Region","(All)","","Style","(All)"],
			["Gender","(All)","","",""]
		],
		"filter_downThenOver7_2wrap":[
			["Region","(All)","","Style","(All)","","Units","(All)","","Cost","(All)"],
			["Gender","(All)","","Ship date","(All)","","Price","(All)","","",""],
		],
		"filter_overThenDown7_2wrap":[
			["Region","(All)","","Gender","(All)"],
			["Style","(All)","","Ship date","(All)"],
			["Units","(All)","","Price","(All)"],
			["Cost","(All)","","",""]
		],
		"data_values1":[
			["Row Labels","Count of Price1","Count of Price2","Min of Price3","Max of Price4","Sum of Price5","Average of Price6","Product of Price7","StdDev of Price8","StdDevp of Price9","Var of Price10","Varp of Price11"],
			["East","6","6","11.04","13.74","73.13","12.18833333","3221490.641","1.028132611","0.938552372","1.057056667","0.880880556"],
			["Boy","3","3","11.04","13","36","12","1716.4992","0.980612054","0.800666389","0.9616","0.641066667"],
			["Girl","3","3","11.27","13.74","37.13","12.37666667","1876.779576","1.254843948","1.024575793","1.574633333","1.049755556"],
			["West","5","5","11.44","13.42","61.03","12.206","268454.7463","0.834973053","0.746822603","0.69718","0.557744"],
			["Boy","3","3","11.44","12.63","36.13","12.04333333","1742.515632","0.595175044","0.485958389","0.354233333","0.236155556"],
			["Girl","2","2","11.48","13.42","24.9","12.45","154.0616","1.371787156","0.97","1.8818","0.9409"],
			["Grand Total","11","11","11.04","13.74","134.16","12.19636364","8.64824E+11","0.898601944","0.856783337","0.807485455","0.734077686"]
		],
		"data_values2": [
			["Row Labels","Count of Price1","Count of Price2","Min of Price3","Max of Price4","Sum of Price5","Average of Price6","Product of Price7","StdDev of Price8","StdDevp of Price9","Var of Price10","Varp of Price11"],
			["East","3","3","11.05","12.12","35.13","11.71","1601.75496","0.577148161","0.4712395","0.3331","0.222066667"],
			["Boy","2","2","11.05","11.96","23.01","11.505","132.158","0.643467171","0.455","0.41405","0.207025"],
			["Girl","1","1","12.12","12.12","12.12","12.12","12.12","#DIV/0!","0","#DIV/0!","0"],
			["West","5","5","11.44","13.42","61.03","12.206","268454.7463","0.834973053","0.746822603","0.69718","0.557744"],
			["Boy","3","3","11.44","12.63","36.13","12.04333333","1742.515632","0.595175044","0.485958389","0.354233333","0.236155556"],
			["Girl","2","2","11.48","13.42","24.9","12.45","154.0616","1.371787156","0.97","1.8818","0.9409"],
			["Grand Total","8","8","11.05","13.42","96.16","12.02","429998721.4","0.747968678","0.699660632","0.559457143","0.489525"]
		],
		"data_values3": [
			["Row Labels","Count of Price1","Count of Price2","Min of Price3","Max of Price4","Sum of Price5","Average of Price6","Product of Price7","StdDev of Price8","StdDevp of Price9","Var of Price10","Varp of Price11"],
			["East","6","3","11.05","12.12","35.13","11.71","1601.75496","0.577148161","0.4712395","0.3331","0.222066667"],
			["Boy","3","2","11.05","11.96","23.01","11.505","132.158","0.643467171","0.455","0.41405","0.207025"],
			["Girl","3","1","12.12","12.12","12.12","12.12","12.12","#DIV/0!","0","#DIV/0!","0"],
			["West","5","5","11.44","13.42","61.03","12.206","268454.7463","0.834973053","0.746822603","0.69718","0.557744"],
			["Boy","3","3","11.44","12.63","36.13","12.04333333","1742.515632","0.595175044","0.485958389","0.354233333","0.236155556"],
			["Girl","2","2","11.48","13.42","24.9","12.45","154.0616","1.371787156","0.97","1.8818","0.9409"],
			["Grand Total","11","8","11.05","13.42","96.16","12.02","429998721.4","0.747968678","0.699660632","0.559457143","0.489525"]
		],
		"data_values4": [
			["Row Labels","Count of Price1","Count of Price2","Min of Price3","Max of Price4","Sum of Price5","Average of Price6","Product of Price7","StdDev of Price8","StdDevp of Price9","Var of Price10","Varp of Price11"],
			["East","6","3","11.05","12.12","35.13","11.71","1601.75496","0.577148161","0.4712395","0.3331","0.222066667"],
			["Boy","3","2","11.05","11.96","23.01","11.505","132.158","0.643467171","0.455","0.41405","0.207025"],
			["Girl","3","1","12.12","12.12","12.12","12.12","12.12","#DIV/0!","0","#DIV/0!","0"],
			["West","5","5","11.44","13.42","61.03","12.206","268454.7463","0.834973053","0.746822603","0.69718","0.557744"],
			["Boy","3","3","11.44","12.63","36.13","12.04333333","1742.515632","0.595175044","0.485958389","0.354233333","0.236155556"],
			["Girl","2","2","11.48","13.42","24.9","12.45","154.0616","1.371787156","0.97","1.8818","0.9409"],
			["Grand Total","11","8","11.05","13.42","96.16","12.02","429998721.4","0.747968678","0.699660632","0.559457143","0.489525"]
		],
		"data_values5": [
			["Row Labels","Count of Price1","Count of Price2","Min of Price3","Max of Price4","Sum of Price5","Average of Price6","Product of Price7","StdDev of Price8","StdDevp of Price9","Var of Price10","Varp of Price11"],
			["East","6","3","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A"],
			["Boy","3","2","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A"],
			["Girl","3","1","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A"],
			["West","5","5","11.44","13.42","61.03","12.206","268454.7463","0.834973053","0.746822603","0.69718","0.557744"],
			["Boy","3","3","11.44","12.63","36.13","12.04333333","1742.515632","0.595175044","0.485958389","0.354233333","0.236155556"],
			["Girl","2","2","11.48","13.42","24.9","12.45","154.0616","1.371787156","0.97","1.8818","0.9409"],
			["Grand Total","11","8","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A","#N/A"]
		],
		"data_values6": [
			["Row Labels","Count of Price1","Count of Price2","Min of Price3","Max of Price4","Sum of Price5","Average of Price6","Product of Price7","StdDev of Price8","StdDevp of Price9","Var of Price10","Varp of Price11"],
			["East","6","6","11.04","13.74","73.13","12.18833333","3221490.641","1.028132611","0.938552372","1.057056667","0.880880556"],
			["Boy","3","3","11.04","13","36","12","1716.4992","0.980612054","0.800666389","0.9616","0.641066667"],
			["Girl","3","3","11.27","13.74","37.13","12.37666667","1876.779576","1.254843948","1.024575793","1.574633333","1.049755556"],
			["West","5","5","11.44","13.42","61.03","12.206","268454.7463","0.834973053","0.746822603","0.69718","0.557744"],
			["Boy","3","3","11.44","12.63","36.13","12.04333333","1742.515632","0.595175044","0.485958389","0.354233333","0.236155556"],
			["Girl","2","2","11.48","13.42","24.9","12.45","154.0616","1.371787156","0.97","1.8818","0.9409"],
			["Grand Total","11","11","11.04","13.74","134.16","12.19636364","8.64824E+11","0.898601944","0.856783337","0.807485455","0.734077686"]
		],
		"data_values7": [
			["Row Labels","Count of Price1","Count of Price2","Min of Price3","Max of Price4","Sum of Price5","Average of Price6","Product of Price7","StdDev of Price8","StdDevp of Price9","Var of Price10","Varp of Price11"],
			["East","1","0","0","0","0","#DIV/0!","0","#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"],
			["Boy","","","","","","","","","","",""],
			["Girl","1","0","0","0","0","#DIV/0!","0","#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"],
			["West","4","3","1","3","6","2","6","1","0.816496581","1","0.666666667"],
			["Boy","2","1","1","1","1","1","1","#DIV/0!","0","#DIV/0!","0"],
			["Girl","2","2","2","3","5","2.5","6","0.707106781","0.5","0.5","0.25"],
			["Grand Total","5","3","1","3","6","2","6","1","0.816496581","1","0.666666667"]
		],
		"fieldSubtotalNone": [
			["Sum of Price","Column Labels","","",""],
			["Row Labels","Fancy","Golf","Tee","Grand Total"],
			["East","","","",""],
			["Boy","11.96","13","11.04","36"],
			["Girl","13.74","12.12","11.27","37.13"],
			["West","","","",""],
			["Boy","12.06","12.63","11.44","36.13"],
			["Girl","","11.48","13.42","24.9"],
			["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"fieldSubtotalAuto": [
			["Sum of Price","Column Labels","","",""],
			["Row Labels","Fancy","Golf","Tee","Grand Total"],
			["East","25.7","25.12","22.31","73.13"],
			["Boy","11.96","13","11.04","36"],
			["Girl","13.74","12.12","11.27","37.13"],
			["West","12.06","24.11","24.86","61.03"],
			["Boy","12.06","12.63","11.44","36.13"],
			["Girl","","11.48","13.42","24.9"],
			["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"fieldSubtotalCustom": [
			["Sum of Price","Column Labels","","",""],
			["Row Labels","Fancy","Golf","Tee","Grand Total"],
			["East","2","2","2","6"],
			["Boy","11.96","13","11.04","36"],
			["Girl","13.74","12.12","11.27","37.13"],
			["West","1","2","2","5"],
			["Boy","12.06","12.63","11.44","36.13"],
			["Girl","","11.48","13.42","24.9"],
			["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"fieldSubtotalAll": [
			["Sum of Price","Column Labels","","",""],
			["Row Labels","Fancy","Golf","Tee","Grand Total"],
			["East","","","",""],
			["Boy","11.96","13","11.04","36"],
			["Girl","13.74","12.12","11.27","37.13"],
			["East Sum","25.7","25.12","22.31","73.13"],
			["East Count","2","2","2","6"],
			["East Average","12.85","12.56","11.155","12.18833333"],
			["East Max","13.74","13","11.27","13.74"],
			["East Min","11.96","12.12","11.04","11.04"],
			["East Product","164.3304","157.56","124.4208","3221490.641"],
			["East Count","2","2","2","6"],
			["East StdDev","1.258650071","0.622253967","0.16263456","1.028132611"],
			["East StdDevp","0.89","0.44","0.115","0.938552372"],
			["East Var","1.5842","0.3872","0.02645","1.057056667"],
			["East Varp","0.7921","0.1936","0.013225","0.880880556"],
			["West","","","",""],
			["Boy","12.06","12.63","11.44","36.13"],
			["Girl","","11.48","13.42","24.9"],
			["West Sum","12.06","24.11","24.86","61.03"],
			["West Count","1","2","2","5"],
			["West Average","12.06","12.055","12.43","12.206"],
			["West Max","12.06","12.63","13.42","13.42"],
			["West Min","12.06","11.48","11.44","11.44"],
			["West Product","12.06","144.9924","153.5248","268454.7463"],
			["West Count","1","2","2","5"],
			["West StdDev","#DIV/0!","0.813172798","1.400071427","0.834973053"],
			["West StdDevp","0","0.575","0.99","0.746822603"],
			["West Var","#DIV/0!","0.66125","1.9602","0.69718"],
			["West Varp","0","0.330625","0.9801","0.557744"],
			["Grand Total","37.76","49.23","47.17","134.16"]
		],
		"field0":[
			["Row Labels","Units",""],
			["East","",""],
			["","10",""],
			["","Girl",""],
			["","Sum of Price",""],
			["","Fancy","13.74"],
			["","Golf","12.12"],
			["","Tee","11.27"],
			["","Sum of Cost",""],
			["","Fancy","13.33"],
			["","Golf","11.95"],
			["","Tee","10.56"],
			["","Girl Sum of Price","37.13"],
			["","Girl Sum of Cost","35.84"],
			["","10 Sum of Price","37.13"],
			["","10 Sum of Cost","35.84"],
			["","12",""],
			["","Boy",""],
			["","Sum of Price",""],
			["","Fancy","11.96"],
			["","Golf","13"],
			["","Tee","11.04"],
			["","Sum of Cost",""],
			["","Fancy","11.74"],
			["","Golf","12.6"],
			["","Tee","10.42"],
			["","Boy Sum of Price","36"],
			["","Boy Sum of Cost","34.76"],
			["","12 Sum of Price","36"],
			["","12 Sum of Cost","34.76"],
			["East Sum of Price","","73.13"],
			["East Sum of Cost","","70.6"],
			["","",""],
			["West","",""],
			["","11",""],
			["","Boy",""],
			["","Sum of Price",""],
			["","Fancy","12.06"],
			["","Golf","12.63"],
			["","Tee","11.44"],
			["","Sum of Cost",""],
			["","Fancy","11.51"],
			["","Golf","11.73"],
			["","Tee","10.94"],
			["","Boy Sum of Price","36.13"],
			["","Boy Sum of Cost","34.18"],
			["","11 Sum of Price","36.13"],
			["","11 Sum of Cost","34.18"],
			["","15",""],
			["","Girl",""],
			["","Sum of Price",""],
			["","Golf","11.48"],
			["","Tee","13.42"],
			["","Sum of Cost",""],
			["","Golf","10.67"],
			["","Tee","13.29"],
			["","Girl Sum of Price","24.9"],
			["","Girl Sum of Cost","23.96"],
			["","15 Sum of Price","24.9"],
			["","15 Sum of Cost","23.96"],
			["West Sum of Price","","61.03"],
			["West Sum of Cost","","58.14"],
			["","",""],
			["Total Sum of Price","","134.16"],
			["Total Sum of Cost","","128.74"]
		],
		"field4":[
			["Row Labels","UnitsCustom","Gender",""],
			["East","","",""],
			["","10","Girl",""],
			["","","Sum of Price",""],
			["","","Fancy","13.74"],
			["","","Golf","12.12"],
			["","","Tee","11.27"],
			["","","Sum of Cost",""],
			["","","Fancy","13.33"],
			["","","Golf","11.95"],
			["","","Tee","10.56"],
			["","","Girl Sum of Price","37.13"],
			["","","Girl Sum of Cost","35.84"],
			["","10 Sum of Price","","37.13"],
			["","10 Sum of Cost","","35.84"],
			["","11","Boy",""],
			["","","Sum of Price",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Sum of Cost",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Boy Sum of Price",""],
			["","","Boy Sum of Cost",""],
			["","","Girl",""],
			["","","Sum of Price",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Sum of Cost",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Girl Sum of Price",""],
			["","","Girl Sum of Cost",""],
			["","11 Sum of Price","",""],
			["","11 Sum of Cost","",""],
			["","12","Boy",""],
			["","","Sum of Price",""],
			["","","Fancy","11.96"],
			["","","Golf","13"],
			["","","Tee","11.04"],
			["","","Sum of Cost",""],
			["","","Fancy","11.74"],
			["","","Golf","12.6"],
			["","","Tee","10.42"],
			["","","Boy Sum of Price","36"],
			["","","Boy Sum of Cost","34.76"],
			["","12 Sum of Price","","36"],
			["","12 Sum of Cost","","34.76"],
			["","15","Boy",""],
			["","","Sum of Price",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Sum of Cost",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Boy Sum of Price",""],
			["","","Boy Sum of Cost",""],
			["","","Girl",""],
			["","","Sum of Price",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Sum of Cost",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Girl Sum of Price",""],
			["","","Girl Sum of Cost",""],
			["","15 Sum of Price","",""],
			["","15 Sum of Cost","",""],
			["East Sum of Price","","","73.13"],
			["East Sum of Cost","","","70.6"],
			["","","",""],
			["West","","",""],
			["","10","Boy",""],
			["","","Sum of Price",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Sum of Cost",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Boy Sum of Price",""],
			["","","Boy Sum of Cost",""],
			["","","Girl",""],
			["","","Sum of Price",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Sum of Cost",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Girl Sum of Price",""],
			["","","Girl Sum of Cost",""],
			["","10 Sum of Price","",""],
			["","10 Sum of Cost","",""],
			["","11","Boy",""],
			["","","Sum of Price",""],
			["","","Fancy","12.06"],
			["","","Golf","12.63"],
			["","","Tee","11.44"],
			["","","Sum of Cost",""],
			["","","Fancy","11.51"],
			["","","Golf","11.73"],
			["","","Tee","10.94"],
			["","","Boy Sum of Price","36.13"],
			["","","Boy Sum of Cost","34.18"],
			["","11 Sum of Price","","36.13"],
			["","11 Sum of Cost","","34.18"],
			["","12","Boy",""],
			["","","Sum of Price",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Sum of Cost",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Boy Sum of Price",""],
			["","","Boy Sum of Cost",""],
			["","","Girl",""],
			["","","Sum of Price",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Sum of Cost",""],
			["","","Fancy",""],
			["","","Golf",""],
			["","","Tee",""],
			["","","Girl Sum of Price",""],
			["","","Girl Sum of Cost",""],
			["","12 Sum of Price","",""],
			["","12 Sum of Cost","",""],
			["","15","Girl",""],
			["","","Sum of Price",""],
			["","","Golf","11.48"],
			["","","Tee","13.42"],
			["","","Sum of Cost",""],
			["","","Golf","10.67"],
			["","","Tee","13.29"],
			["","","Girl Sum of Price","24.9"],
			["","","Girl Sum of Cost","23.96"],
			["","15 Sum of Price","","24.9"],
			["","15 Sum of Cost","","23.96"],
			["West Sum of Price","","","61.03"],
			["West Sum of Cost","","","58.14"],
			["","","",""],
			["Total Sum of Price","","","134.16"],
			["Total Sum of Cost","","","128.74"]
		],
		"longHeader":[
			["Sum of Price","Gender","Style"],
			["","Boy",""],
			["Region","Tee",""],
			["East","11.04",""]
		],
		"headerRename":[
			["1","(All)","",""],
			["qwe","(All)","",""],
			["TRUE","(All)","",""],
			["#DIV/0!","(All)","",""],
			["1.234567891","(All)","",""],
			["12345678912","(All)","",""],
			["1.00","(All)","",""],
			["$2.00","(All)","",""],
			["Tuesday, January 03, 1900","(All)","",""],
			["400.00%","(All)","",""],
			["5.00E+00","(All)","",""],
			["qwerty","(All)","",""],
			["12","(All)","",""],
			["qwe2","(All)","",""],
			["","","",""],
			["Sum of 1","Sum of 1_2","Sum of qwe","Sum of qwe2"],
			["1","1","1","1"]
		],
		"addField":[
			["Row Labels","Sum of Ship date","Sum of Units","Sum of Price","Sum of Cost","Count of Region","Count of Gender","Count of Style"],
			["East","230298","66","73.13","70.6","6","6","6"],
			["Boy","115149","36","36","34.76","3","3","3"],
			["Fancy","38383","12","11.96","11.74","1","1","1"],
			["Golf","38383","12","13","12.6","1","1","1"],
			["Tee","38383","12","11.04","10.42","1","1","1"],
			["Girl","115149","30","37.13","35.84","3","3","3"],
			["Fancy","38383","10","13.74","13.33","1","1","1"],
			["Golf","38383","10","12.12","11.95","1","1","1"],
			["Tee","38383","10","11.27","10.56","1","1","1"],
			["West","191915","63","61.03","58.14","5","5","5"],
			["Boy","115149","33","36.13","34.18","3","3","3"],
			["Fancy","38383","11","12.06","11.51","1","1","1"],
			["Golf","38383","11","12.63","11.73","1","1","1"],
			["Tee","38383","11","11.44","10.94","1","1","1"],
			["Girl","76766","30","24.9","23.96","2","2","2"],
			["Golf","38383","15","11.48","10.67","1","1","1"],
			["Tee","38383","15","13.42","13.29","1","1","1"],
			["Grand Total","422213","129","134.16","128.74","11","11","11"]
		],
		"moveDataField":[
			["Price","(All)","","","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","","","","",""],
			["","Column Labels","","","","","","","","","","","","","","","","",""],
			["","Sum of Cost","","Count of Region","","Count of Gender","","Count of Style","","Sum of Price2","","Sum of Units","","Total Sum of Cost","Total Count of Region","Total Count of Gender","Total Count of Style","Total Sum of Price2","Total Sum of Units"],
			["Row Labels","East","West","East","West","East","West","East","West","East","West","East","West","","","","","",""],
			["Boy","34.76","34.18","3","3","3","3","3","3","36","36.13","36","33","68.94","6","6","6","72.13","69"],
			["Fancy","11.74","11.51","1","1","1","1","1","1","11.96","12.06","12","11","23.25","2","2","2","24.02","23"],
			["38383","11.74","11.51","1","1","1","1","1","1","11.96","12.06","12","11","23.25","2","2","2","24.02","23"],
			["Golf","12.6","11.73","1","1","1","1","1","1","13","12.63","12","11","24.33","2","2","2","25.63","23"],
			["38383","12.6","11.73","1","1","1","1","1","1","13","12.63","12","11","24.33","2","2","2","25.63","23"],
			["Tee","10.42","10.94","1","1","1","1","1","1","11.04","11.44","12","11","21.36","2","2","2","22.48","23"],
			["38383","10.42","10.94","1","1","1","1","1","1","11.04","11.44","12","11","21.36","2","2","2","22.48","23"],
			["Girl","35.84","23.96","3","2","3","2","3","2","37.13","24.9","30","30","59.8","5","5","5","62.03","60"],
			["Fancy","13.33","","1","","1","","1","","13.74","","10","","13.33","1","1","1","13.74","10"],
			["38383","13.33","","1","","1","","1","","13.74","","10","","13.33","1","1","1","13.74","10"],
			["Golf","11.95","10.67","1","1","1","1","1","1","12.12","11.48","10","15","22.62","2","2","2","23.6","25"],
			["38383","11.95","10.67","1","1","1","1","1","1","12.12","11.48","10","15","22.62","2","2","2","23.6","25"],
			["Tee","10.56","13.29","1","1","1","1","1","1","11.27","13.42","10","15","23.85","2","2","2","24.69","25"],
			["38383","10.56","13.29","1","1","1","1","1","1","11.27","13.42","10","15","23.85","2","2","2","24.69","25"],
			["Grand Total","70.6","58.14","6","5","6","5","6","5","73.13","61.03","66","63","128.74","11","11","11","134.16","129"]
		],
		"moveToField":[
			["Style","(All)","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","Column Labels","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","Sum of Cost","","","","","","","","","","","Count of Region","","","","","","","","","","","Count of Gender","","","","","","","","","","","Count of Style","","","","","","","","","","","Sum of Price2","","","","","","","","","","","Count of Region2","","","","","","","","","","","Total Sum of Cost","Total Count of Region","Total Count of Gender","Total Count of Style","Total Sum of Price2","Total Count of Region2"],
			["Row Labels","11.04","11.27","11.44","11.48","11.96","12.06","12.12","12.63","13","13.42","13.74","11.04","11.27","11.44","11.48","11.96","12.06","12.12","12.63","13","13.42","13.74","11.04","11.27","11.44","11.48","11.96","12.06","12.12","12.63","13","13.42","13.74","11.04","11.27","11.44","11.48","11.96","12.06","12.12","12.63","13","13.42","13.74","11.04","11.27","11.44","11.48","11.96","12.06","12.12","12.63","13","13.42","13.74","11.04","11.27","11.44","11.48","11.96","12.06","12.12","12.63","13","13.42","13.74","","","","","",""],
			["Boy","10.42","","10.94","","11.74","11.51","","11.73","12.6","","","1","","1","","1","1","","1","1","","","1","","1","","1","1","","1","1","","","1","","1","","1","1","","1","1","","","11.04","","11.44","","11.96","12.06","","12.63","13","","","1","","1","","1","1","","1","1","","","68.94","6","6","6","72.13","6"],
			["38383","10.42","","10.94","","11.74","11.51","","11.73","12.6","","","1","","1","","1","1","","1","1","","","1","","1","","1","1","","1","1","","","1","","1","","1","1","","1","1","","","11.04","","11.44","","11.96","12.06","","12.63","13","","","1","","1","","1","1","","1","1","","","68.94","6","6","6","72.13","6"],
			["11","","","10.94","","","11.51","","11.73","","","","","","1","","","1","","1","","","","","","1","","","1","","1","","","","","","1","","","1","","1","","","","","","11.44","","","12.06","","12.63","","","","","","1","","","1","","1","","","","34.18","3","3","3","36.13","3"],
			["12","10.42","","","","11.74","","","","12.6","","","1","","","","1","","","","1","","","1","","","","1","","","","1","","","1","","","","1","","","","1","","","11.04","","","","11.96","","","","13","","","1","","","","1","","","","1","","","34.76","3","3","3","36","3"],
			["Girl","","10.56","","10.67","","","11.95","","","13.29","13.33","","1","","1","","","1","","","1","1","","1","","1","","","1","","","1","1","","1","","1","","","1","","","1","1","","11.27","","11.48","","","12.12","","","13.42","13.74","","1","","1","","","1","","","1","1","59.8","5","5","5","62.03","5"],
			["38383","","10.56","","10.67","","","11.95","","","13.29","13.33","","1","","1","","","1","","","1","1","","1","","1","","","1","","","1","1","","1","","1","","","1","","","1","1","","11.27","","11.48","","","12.12","","","13.42","13.74","","1","","1","","","1","","","1","1","59.8","5","5","5","62.03","5"],
			["10","","10.56","","","","","11.95","","","","13.33","","1","","","","","1","","","","1","","1","","","","","1","","","","1","","1","","","","","1","","","","1","","11.27","","","","","12.12","","","","13.74","","1","","","","","1","","","","1","35.84","3","3","3","37.13","3"],
			["15","","","","10.67","","","","","","13.29","","","","","1","","","","","","1","","","","","1","","","","","","1","","","","","1","","","","","","1","","","","","11.48","","","","","","13.42","","","","","1","","","","","","1","","23.96","2","2","2","24.9","2"],
			["Grand Total","10.42","10.56","10.94","10.67","11.74","11.51","11.95","11.73","12.6","13.29","13.33","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","1","11.04","11.27","11.44","11.48","11.96","12.06","12.12","12.63","13","13.42","13.74","1","1","1","1","1","1","1","1","1","1","1","128.74","11","11","11","134.16","11"]
		],
		"moveField":[
			["Price","(All)","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["Style","(All)","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["Cost","(All)","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","Column Labels","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","10","","","","","10 Sum of Price","10 Count of Region","10 Count of Gender","10 Sum of Cost","10 Sum of Price2","11","","","","","11 Sum of Price","11 Count of Region","11 Count of Gender","11 Sum of Cost","11 Sum of Price2","12","","","","","12 Sum of Price","12 Count of Region","12 Count of Gender","12 Sum of Cost","12 Sum of Price2","15","","","","","15 Sum of Price","15 Count of Region","15 Count of Gender","15 Sum of Cost","15 Sum of Price2","Total Sum of Price","Total Count of Region","Total Count of Gender","Total Sum of Cost","Total Sum of Price2"],
			["","Sum of Price","Count of Region","Count of Gender","Sum of Cost","Sum of Price2","","","","","","Sum of Price","Count of Region","Count of Gender","Sum of Cost","Sum of Price2","","","","","","Sum of Price","Count of Region","Count of Gender","Sum of Cost","Sum of Price2","","","","","","Sum of Price","Count of Region","Count of Gender","Sum of Cost","Sum of Price2","","","","","","","","","",""],
			["Row Labels","38383","38383","38383","38383","38383","","","","","","38383","38383","38383","38383","38383","","","","","","38383","38383","38383","38383","38383","","","","","","38383","38383","38383","38383","38383","","","","","","","","","",""],
			["Boy","","","","","","","","","","","36.13","3","3","34.18","36.13","36.13","3","3","34.18","36.13","36","3","3","34.76","36","36","3","3","34.76","36","","","","","","","","","","","72.13","6","6","68.94","72.13"],
			["East","","","","","","","","","","","","","","","","","","","","","36","3","3","34.76","36","36","3","3","34.76","36","","","","","","","","","","","36","3","3","34.76","36"],
			["West","","","","","","","","","","","36.13","3","3","34.18","36.13","36.13","3","3","34.18","36.13","","","","","","","","","","","","","","","","","","","","","36.13","3","3","34.18","36.13"],
			["Girl","37.13","3","3","35.84","37.13","37.13","3","3","35.84","37.13","","","","","","","","","","","","","","","","","","","","","24.9","2","2","23.96","24.9","24.9","2","2","23.96","24.9","62.03","5","5","59.8","62.03"],
			["East","37.13","3","3","35.84","37.13","37.13","3","3","35.84","37.13","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","37.13","3","3","35.84","37.13"],
			["West","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","24.9","2","2","23.96","24.9","24.9","2","2","23.96","24.9","24.9","2","2","23.96","24.9"],
			["Grand Total","37.13","3","3","35.84","37.13","37.13","3","3","35.84","37.13","36.13","3","3","34.18","36.13","36.13","3","3","34.18","36.13","36","3","3","34.76","36","36","3","3","34.76","36","24.9","2","2","23.96","24.9","24.9","2","2","23.96","24.9","134.16","11","11","128.74","134.16"]
		],
		"removeDataField":[
			["Price","(All)","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["Style","(All)","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","Column Labels","","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","10","","","10 Count of Region","10 Count of Gender","10 Sum of Price2","11","","","11 Count of Region","11 Count of Gender","11 Sum of Price2","12","","","12 Count of Region","12 Count of Gender","12 Sum of Price2","15","","","15 Count of Region","15 Count of Gender","15 Sum of Price2","Total Count of Region","Total Count of Gender","Total Sum of Price2"],
			["","Count of Region","Count of Gender","Sum of Price2","","","","Count of Region","Count of Gender","Sum of Price2","","","","Count of Region","Count of Gender","Sum of Price2","","","","Count of Region","Count of Gender","Sum of Price2","","","","","",""],
			["Row Labels","38383","38383","38383","","","","38383","38383","38383","","","","38383","38383","38383","","","","38383","38383","38383","","","","","",""],
			["Boy","","","","","","","3","3","36.13","3","3","36.13","3","3","36","3","3","36","","","","","","","6","6","72.13"],
			["Girl","3","3","37.13","3","3","37.13","","","","","","","","","","","","","2","2","24.9","2","2","24.9","5","5","62.03"],
			["Grand Total","3","3","37.13","3","3","37.13","3","3","36.13","3","3","36.13","3","3","36","3","3","36","2","2","24.9","2","2","24.9","11","11","134.16"]
		],
		"dataOnCols":[
			["Row Labels","Sum of Price","Sum of Cost"],
			["East","73.13","70.6"],
			["Boy","36","34.76"],
			["Fancy","11.96","11.74"],
			["Golf","13","12.6"],
			["Tee","11.04","10.42"],
			["Girl","37.13","35.84"],
			["Fancy","13.74","13.33"],
			["Golf","12.12","11.95"],
			["Tee","11.27","10.56"],
			["West","61.03","58.14"],
			["Boy","36.13","34.18"],
			["Fancy","12.06","11.51"],
			["Golf","12.63","11.73"],
			["Tee","11.44","10.94"],
			["Girl","24.9","23.96"],
			["Golf","11.48","10.67"],
			["Tee","13.42","13.29"],
			["Grand Total","134.16","128.74"]
		],
		"dataOnRows":[
			["Row Labels",""],
			["East",""],
			["Boy",""],
			["Fancy",""],
			["Sum of Price","11.96"],
			["Sum of Cost","11.74"],
			["Golf",""],
			["Sum of Price","13"],
			["Sum of Cost","12.6"],
			["Tee",""],
			["Sum of Price","11.04"],
			["Sum of Cost","10.42"],
			["Boy Sum of Price","36"],
			["Boy Sum of Cost","34.76"],
			["Girl",""],
			["Fancy",""],
			["Sum of Price","13.74"],
			["Sum of Cost","13.33"],
			["Golf",""],
			["Sum of Price","12.12"],
			["Sum of Cost","11.95"],
			["Tee",""],
			["Sum of Price","11.27"],
			["Sum of Cost","10.56"],
			["Girl Sum of Price","37.13"],
			["Girl Sum of Cost","35.84"],
			["East Sum of Price","73.13"],
			["East Sum of Cost","70.6"],
			["West",""],
			["Boy",""],
			["Fancy",""],
			["Sum of Price","12.06"],
			["Sum of Cost","11.51"],
			["Golf",""],
			["Sum of Price","12.63"],
			["Sum of Cost","11.73"],
			["Tee",""],
			["Sum of Price","11.44"],
			["Sum of Cost","10.94"],
			["Boy Sum of Price","36.13"],
			["Boy Sum of Cost","34.18"],
			["Girl",""],
			["Golf",""],
			["Sum of Price","11.48"],
			["Sum of Cost","10.67"],
			["Tee",""],
			["Sum of Price","13.42"],
			["Sum of Cost","13.29"],
			["Girl Sum of Price","24.9"],
			["Girl Sum of Cost","23.96"],
			["West Sum of Price","61.03"],
			["West Sum of Cost","58.14"],
			["Total Sum of Price","134.16"],
			["Total Sum of Cost","128.74"]
		],
		"dataOnRowsValues":[
			["Row Labels"],
			["East"],
			["Boy"],
			["Fancy"],
			["Golf"],
			["Tee"],
			["Girl"],
			["Fancy"],
			["Golf"],
			["Tee"],
			["West"],
			["Boy"],
			["Fancy"],
			["Golf"],
			["Tee"],
			["Girl"],
			["Golf"],
			["Tee"],
			["Grand Total"]
		],
		"dataPosition":[
			["Row Labels",""],
			["East",""],
			["Boy",""],
			["Sum of Price",""],
			["Fancy","11.96"],
			["Golf","13"],
			["Tee","11.04"],
			["Sum of Cost",""],
			["Fancy","11.74"],
			["Golf","12.6"],
			["Tee","10.42"],
			["Boy Sum of Price","36"],
			["Boy Sum of Cost","34.76"],
			["Girl",""],
			["Sum of Price",""],
			["Fancy","13.74"],
			["Golf","12.12"],
			["Tee","11.27"],
			["Sum of Cost",""],
			["Fancy","13.33"],
			["Golf","11.95"],
			["Tee","10.56"],
			["Girl Sum of Price","37.13"],
			["Girl Sum of Cost","35.84"],
			["East Sum of Price","73.13"],
			["East Sum of Cost","70.6"],
			["West",""],
			["Boy",""],
			["Sum of Price",""],
			["Fancy","12.06"],
			["Golf","12.63"],
			["Tee","11.44"],
			["Sum of Cost",""],
			["Fancy","11.51"],
			["Golf","11.73"],
			["Tee","10.94"],
			["Boy Sum of Price","36.13"],
			["Boy Sum of Cost","34.18"],
			["Girl",""],
			["Sum of Price",""],
			["Golf","11.48"],
			["Tee","13.42"],
			["Sum of Cost",""],
			["Golf","10.67"],
			["Tee","13.29"],
			["Girl Sum of Price","24.9"],
			["Girl Sum of Cost","23.96"],
			["West Sum of Price","61.03"],
			["West Sum of Cost","58.14"],
			["Total Sum of Price","134.16"],
			["Total Sum of Cost","128.74"]
		],
		"refreshFieldSettings":[
			["RenamedUnits","(All)","","","","","","","","","","","",""],
			["Cost","(All)","","","","","","","","","","","",""],
			["","","","","","","","","","","","","",""],
			["","","Column Labels","","","","","","","","","","",""],
			["","","Fancy","","","Golf","","","Tee","","","Total Average of Price","Total Count of Region","Total Sum of Cost"],
			["","","38383","","","38383","","","38383","","","","",""],
			["Row Labels","Gender","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","","",""],
			["East","Boy","11.96","1","11.74","13","1","12.6","11.04","1","10.42","12","3","34.76"],
			["","Girl","13.74","1","13.33","12.12","1","11.95","11.27","1","10.56","12.37666667","3","35.84"],
			["East Average","","12.85","#DIV/0!","12.535","12.56","#DIV/0!","12.275","11.155","#DIV/0!","10.49","12.18833333","#DIV/0!","11.76666667"],
			["West","Boy","12.06","1","11.51","12.63","1","11.73","11.44","1","10.94","12.04333333","3","34.18"],
			["","Girl","","","","11.48","1","10.67","13.42","1","13.29","12.45","2","23.96"],
			["West Average","","12.06","#DIV/0!","11.51","12.055","#DIV/0!","11.2","12.43","#DIV/0!","12.115","12.206","#DIV/0!","11.628"],
			["Grand Total","","12.58666667","3","36.58","12.3075","4","46.95","11.7925","4","45.21","12.19636364","11","128.74"]
		],
		"refreshRecords":[
			["RenamedUnits","(All)","","","","","","","","","","","","","","","","","","","","","","","",""],
			["Cost","(All)","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","Column Labels","","","","","","","","","","","","","","","","","","","","","","",""],
			["","","Fancy","","","","","","Golf","","","","","","Tee","","","WWW","","","BBB","","","Total Average of Price","Total Count of Region","Total Sum of Cost"],
			["","","38383","","","10","","","38383","","","100000","","","38383","","","38383","","","38383","","","","",""],
			["Row Labels","Gender","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","","",""],
			["East","Boy","","","","11.96","1","11.74","","","","13","1","12.6","11.04","1","10.42","","","","","","","12","3","34.76"],
			["","Girl","13.74","1","13.33","","","","","","","","","","","","","","","","","","","13.74","1","13.33"],
			["","Dog","","","","","","","12.12","1","11.95","","","","","","","","","","","","","12.12","1","11.95"],
			["East Average","","13.74","#DIV/0!","13.33","11.96","#DIV/0!","11.74","12.12","#DIV/0!","11.95","13","#DIV/0!","12.6","11.04","#DIV/0!","10.42","","","","","","","12.372","#DIV/0!","12.008"],
			["West","Boy","12.06","1","11.51","","","","12.63","1","11.73","","","","","","","11.44","1","10.94","","","","12.04333333","3","34.18"],
			["","Girl","","","","","","","11.48","1","10.67","","","","","","","","","","13.42","1","13.29","12.45","2","23.96"],
			["West Average","","12.06","#DIV/0!","11.51","","","","12.055","#DIV/0!","11.2","","","","","","","11.44","#DIV/0!","10.94","13.42","#DIV/0!","13.29","12.206","#DIV/0!","11.628"],
			["North","Girl","","","","","","","","","","","","","11.27","1","10.56","","","","","","","11.27","1","10.56"],
			["North Average","","","","","","","","","","","","","","11.27","#DIV/0!","10.56","","","","","","","11.27","#DIV/0!","10.56"],
			["Grand Total","","12.9","2","24.84","11.96","1","11.74","12.07666667","3","34.35","13","1","12.6","11.155","2","20.98","11.44","1","10.94","13.42","1","13.29","12.19636364","11","128.74"]
		],
		"refreshStructure":[
			["Cost","(All)","","","","","","","","",""],
			["","","","","","","","","","",""],
			["","","Column Labels","","","","","","","",""],
			["","","East","","","West","","","Total Average of Price","Total Count of Region","Total Sum of Cost"],
			["Row Labels","Gender","Average of Price","Count of Region","Sum of Cost","Average of Price","Count of Region","Sum of Cost","","",""],
			["11.04","Boy","#DIV/0!","1","1","","","","#DIV/0!","1","1"],
			["11.04 Average","","#DIV/0!","11.04","1","","","","#DIV/0!","11.04","1"],
			["13","Boy","#DIV/0!","1","2","","","","#DIV/0!","1","2"],
			["13 Average","","#DIV/0!","13","2","","","","#DIV/0!","13","2"],
			["11.96","Boy","#DIV/0!","1","3","","","","#DIV/0!","1","3"],
			["11.96 Average","","#DIV/0!","11.96","3","","","","#DIV/0!","11.96","3"],
			["11.27","Girl","#DIV/0!","1","4","","","","#DIV/0!","1","4"],
			["11.27 Average","","#DIV/0!","11.27","4","","","","#DIV/0!","11.27","4"],
			["12.12","Girl","#DIV/0!","1","5","","","","#DIV/0!","1","5"],
			["12.12 Average","","#DIV/0!","12.12","5","","","","#DIV/0!","12.12","5"],
			["13.74","Girl","#DIV/0!","1","6","","","","#DIV/0!","1","6"],
			["13.74 Average","","#DIV/0!","13.74","6","","","","#DIV/0!","13.74","6"],
			["11.44","Boy","","","","#DIV/0!","1","7","#DIV/0!","1","7"],
			["11.44 Average","","","","","#DIV/0!","11.44","7","#DIV/0!","11.44","7"],
			["12.63","Boy","","","","#DIV/0!","1","8","#DIV/0!","1","8"],
			["12.63 Average","","","","","#DIV/0!","12.63","8","#DIV/0!","12.63","8"],
			["12.06","Boy","","","","#DIV/0!","1","9","#DIV/0!","1","9"],
			["12.06 Average","","","","","#DIV/0!","12.06","9","#DIV/0!","12.06","9"],
			["13.42","Girl","","","","#DIV/0!","1","10","#DIV/0!","1","10"],
			["13.42 Average","","","","","#DIV/0!","13.42","10","#DIV/0!","13.42","10"],
			["11.48","Girl","","","","#DIV/0!","1","11","#DIV/0!","1","11"],
			["11.48 Average","","","","","#DIV/0!","11.48","11","#DIV/0!","11.48","11"],
			["Grand Total","","#DIV/0!","6","21","#DIV/0!","5","45","#DIV/0!","11","66"]
		],
		"valueFilterOrder1":[
			["Sum of Price","","Region","Units","","",""],
			["","","West","","","West Total","Grand Total"],
			["Gender","Ship date","4","5","6","",""],
			["Boy","2","","","","",""],
			["","4","7","20","","27","27"],
			["","5","","20","","20","20"],
			["Boy Total","","7","40","","47","47"],
			["Grand Total","","7","40","","47","47"]
		],
		"valueFilterOrder2":[
			["Sum of Price","","Region","Units","","","","","","",""],
			["","","East","","","East Total","West","","","West Total","Grand Total"],
			["Gender","Ship date","2","3","4","","4","5","6","",""],
			["Boy","2","3","","","3","","","","","3"],
			["","4","","","","","7","20","","27","27"],
			["","5","","","","","","20","","20","20"],
			["Boy Total","","3","","","3","7","40","","47","50"],
			["Girl","2","","4","","4","","","","","4"],
			["","3","","5","6","11","","","","","11"],
			["","6","","","","","","","20","20","20"],
			["Girl Total","","","9","6","15","","","20","20","35"],
			["Grand Total","","3","9","6","18","7","40","20","67","85"]
		],
		"bug-46141-row": [
			["Region","Gender","Style","Sum of Price"],
			["East","Girl","Fancy","13.74"],
			["","Girl Total","","13.74"],
			["East Total","","","13.74"],
			["Grand Total","","","13.74"]
		],
		"bug-46141-col" : [
			["","Region","Gender","Style",""],
			["","East","","East Total","Grand Total"],
			["","Girl","Girl Total","",""],
			["","Fancy","","",""],
			["Sum of Price","13.74","13.74","13.74","13.74"]
		],
		"top10":[
			["","","","Gender","Values","","","",""],
			["","","","Boy","","Girl","","Total Sum of Price","Total Sum of Cost"],
			["Region","Style","Units","Sum of Price","Sum of Cost","Sum of Price","Sum of Cost","",""],
			["West","Fancy","11","12.06","11.51","","","12.06","11.51"],
			["","Fancy Total","","12.06","11.51","","","12.06","11.51"],
			["West Total","","","12.06","11.51","","","12.06","11.51"],
			["Grand Total","","","12.06","11.51","","","12.06","11.51"]
		],
		"label1":[
			["Units","Price","Cost","Sum of Ship date"],
			["10","13.74","13.33","38383"],
			["","13.74 Total","","38383"],
			["10 Total","","","38383"],
			["12","13","12.6","38383"],
			["","13 Total","","38383"],
			["12 Total","","","38383"],
			["15","13.42","13.29","38383"],
			["","13.42 Total","","38383"],
			["15 Total","","","38383"],
			["Grand Total","","","115149"]
		],
		"reIndexStart":[
			["","","Region","Gender","Values","","","","","","","","","","","","",""],
			["","","West","","","","","","","","West Sum of Price","West Sum of Cost","West Sum of Price2","West Sum of Cost2","Total Sum of Price","Total Sum of Cost","Total Sum of Price2","Total Sum of Cost2"],
			["","","Boy","","","","Girl","","","","","","","","","","",""],
			["Style","Units","Sum of Price","Sum of Cost","Sum of Price2","Sum of Cost2","Sum of Price","Sum of Cost","Sum of Price2","Sum of Cost2","","","","","","","",""],
			["Golf","11","12.63","11.73","12.63","11.73","","","","","12.63","11.73","12.63","11.73","12.63","11.73","12.63","11.73"],
			["","15","","","","","11.48","10.67","11.48","10.67","11.48","10.67","11.48","10.67","11.48","10.67","11.48","10.67"],
			["","10","","","","","","","","","","","","","","","",""],
			["","12","","","","","","","","","","","","","","","",""],
			["Golf Total","","12.63","11.73","12.63","11.73","11.48","10.67","11.48","10.67","24.11","22.4","24.11","22.4","24.11","22.4","24.11","22.4"],
			["Tee","15","","","","","13.42","13.29","13.42","13.29","13.42","13.29","13.42","13.29","13.42","13.29","13.42","13.29"],
			["","11","11.44","10.94","11.44","10.94","","","","","11.44","10.94","11.44","10.94","11.44","10.94","11.44","10.94"],
			["","10","","","","","","","","","","","","","","","",""],
			["","12","","","","","","","","","","","","","","","",""],
			["Tee Total","","11.44","10.94","11.44","10.94","13.42","13.29","13.42","13.29","24.86","24.23","24.86","24.23","24.86","24.23","24.86","24.23"],
			["Grand Total","","24.07","22.67","24.07","22.67","24.9","23.96","24.9","23.96","48.97","46.63","48.97","46.63","48.97","46.63","48.97","46.63"]
		],
		"reIndexMove":[
			["","","Region","Gender","Values","","","","","","","","","","","","",""],
			["","","West","","","","","","","","West Sum of Cost2","West Sum of Price","West Sum of Cost","West Sum of Price2","Total Sum of Cost2","Total Sum of Price","Total Sum of Cost","Total Sum of Price2"],
			["","","Boy","","","","Girl","","","","","","","","","","",""],
			["Style","Units","Sum of Cost2","Sum of Price","Sum of Cost","Sum of Price2","Sum of Cost2","Sum of Price","Sum of Cost","Sum of Price2","","","","","","","",""],
			["Golf","11","11.73","12.63","11.73","12.63","","","","","11.73","12.63","11.73","12.63","11.73","12.63","11.73","12.63"],
			["","15","","","","","10.67","11.48","10.67","11.48","10.67","11.48","10.67","11.48","10.67","11.48","10.67","11.48"],
			["","10","","","","","","","","","","","","","","","",""],
			["","12","","","","","","","","","","","","","","","",""],
			["Golf Total","","11.73","12.63","11.73","12.63","10.67","11.48","10.67","11.48","22.4","24.11","22.4","24.11","22.4","24.11","22.4","24.11"],
			["Tee","15","","","","","13.29","13.42","13.29","13.42","13.29","13.42","13.29","13.42","13.29","13.42","13.29","13.42"],
			["","11","10.94","11.44","10.94","11.44","","","","","10.94","11.44","10.94","11.44","10.94","11.44","10.94","11.44"],
			["","10","","","","","","","","","","","","","","","",""],
			["","12","","","","","","","","","","","","","","","",""],
			["Tee Total","","10.94","11.44","10.94","11.44","13.29","13.42","13.29","13.42","24.23","24.86","24.23","24.86","24.23","24.86","24.23","24.86"],
			["Grand Total","","22.67","24.07","22.67","24.07","23.96","24.9","23.96","24.9","46.63","48.97","46.63","48.97","46.63","48.97","46.63","48.97"]
		],
		"reIndexDelete":[
			["","","Region","Gender","Values","","","","","","","","","","",""],
			["","","East","","","","East Sum of Price","East Sum of Price2","West","","","","West Sum of Price","West Sum of Price2","Total Sum of Price","Total Sum of Price2"],
			["","","Girl","","Boy","","","","Boy","","Girl","","","","",""],
			["Style","Units","Sum of Price","Sum of Price2","Sum of Price","Sum of Price2","","","Sum of Price","Sum of Price2","Sum of Price","Sum of Price2","","","",""],
			["Golf","10","12.12","12.12","","","12.12","12.12","","","","","","","12.12","12.12"],
			["","11","","","","","","","12.63","12.63","","","12.63","12.63","12.63","12.63"],
			["","12","","","13","13","13","13","","","","","","","13","13"],
			["","15","","","","","","","","","11.48","11.48","11.48","11.48","11.48","11.48"],
			["Golf Total","","12.12","12.12","13","13","25.12","25.12","12.63","12.63","11.48","11.48","24.11","24.11","49.23","49.23"],
			["Tee","10","11.27","11.27","","","11.27","11.27","","","","","","","11.27","11.27"],
			["","11","","","","","","","11.44","11.44","","","11.44","11.44","11.44","11.44"],
			["","12","","","11.04","11.04","11.04","11.04","","","","","","","11.04","11.04"],
			["","15","","","","","","","","","13.42","13.42","13.42","13.42","13.42","13.42"],
			["Tee Total","","11.27","11.27","11.04","11.04","22.31","22.31","11.44","11.44","13.42","13.42","24.86","24.86","47.17","47.17"],
			["Grand Total","","23.39","23.39","24.04","24.04","47.43","47.43","24.07","24.07","24.9","24.9","48.97","48.97","96.4","96.4"]
		],
		"reIndexAdd":[
			["","","Region","Gender","Values","","","","","","","","","","","","","","","","","",""],
			["","","East","","","","","","East Sum of Price","East Sum of Cost","East Sum of Price2","West","","","","","","West Sum of Price","West Sum of Cost","West Sum of Price2","Total Sum of Price","Total Sum of Cost","Total Sum of Price2"],
			["","","Girl","","","Boy","","","","","","Boy","","","Girl","","","","","","","",""],
			["Style","Units","Sum of Price","Sum of Cost","Sum of Price2","Sum of Price","Sum of Cost","Sum of Price2","","","","Sum of Price","Sum of Cost","Sum of Price2","Sum of Price","Sum of Cost","Sum of Price2","","","","","",""],
			["Golf","10","12.12","11.95","12.12","","","","12.12","11.95","12.12","","","","","","","","","","12.12","11.95","12.12"],
			["","11","","","","","","","","","","12.63","11.73","12.63","","","","12.63","11.73","12.63","12.63","11.73","12.63"],
			["","12","","","","13","12.6","13","13","12.6","13","","","","","","","","","","13","12.6","13"],
			["","15","","","","","","","","","","","","","11.48","10.67","11.48","11.48","10.67","11.48","11.48","10.67","11.48"],
			["Golf Total","","12.12","11.95","12.12","13","12.6","13","25.12","24.55","25.12","12.63","11.73","12.63","11.48","10.67","11.48","24.11","22.4","24.11","49.23","46.95","49.23"],
			["Tee","10","11.27","10.56","11.27","","","","11.27","10.56","11.27","","","","","","","","","","11.27","10.56","11.27"],
			["","11","","","","","","","","","","11.44","10.94","11.44","","","","11.44","10.94","11.44","11.44","10.94","11.44"],
			["","12","","","","11.04","10.42","11.04","11.04","10.42","11.04","","","","","","","","","","11.04","10.42","11.04"],
			["","15","","","","","","","","","","","","","13.42","13.29","13.42","13.42","13.29","13.42","13.42","13.29","13.42"],
			["Tee Total","","11.27","10.56","11.27","11.04","10.42","11.04","22.31","20.98","22.31","11.44","10.94","11.44","13.42","13.29","13.42","24.86","24.23","24.86","47.17","45.21","47.17"],
			["Grand Total","","23.39","22.51","23.39","24.04","23.02","24.04","47.43","45.53","47.43","24.07","22.67","24.07","24.9","23.96","24.9","48.97","46.63","48.97","96.4","92.16","96.4"]
		],
		"numFormat" : [
			["Units","1000.0%","","","","","","","","","","","","","","","","","","","","",""],
			["","","","","","","","","","","","","","","","","","","","","","",""],
			["Sum of Price","","","","Price","hasBlank","Mix","MixFormat","bool","error","","","","","","","","","","","","",""],
			["","","","","$11.3","","","","","$11.3 Total","$12.1","","","","","$12.1 Total","$13.7","","","","","$13.7 Total","Grand Total"],
			["","","","","(blank)","","","","(blank) Total","","11.95","","","","11.95 Total","","13.33","","","","13.33 Total","",""],
			["","","","","qwe","","","qwe Total","","","11.95","","","11.95 Total","","","13.33","","","13.33 Total","","",""],
			["","","","","11.27","","11.27 Total","","","","11.95","","11.95 Total","","","","13.33","","13.33 Total","","","",""],
			["","","","","TRUE","TRUE Total","","","","","FALSE","FALSE Total","","","","","FALSE","FALSE Total","","","","",""],
			["Text","Date","Units2","Units3","#DIV/0!","","","","","","#DIV/0!","","","","","","#DIV/0!","","","","","",""],
			["Fancy","02/05/05","$ 10.0","1.0E+01","","","","","","","","","","","","","13.74","13.74","13.74","13.74","13.74","13.74","13.74"],
			["",""," $10.0  Total","","","","","","","","","","","","","","13.74","13.74","13.74","13.74","13.74","13.74","13.74"],
			["","02/05/05 Total","","","","","","","","","","","","","","","13.74","13.74","13.74","13.74","13.74","13.74","13.74"],
			["Fancy Total","","","","","","","","","","","","","","","","13.74","13.74","13.74","13.74","13.74","13.74","13.74"],
			["Golf","02/04/05","$ 10.0","1.0E+01","","","","","","","12.12","12.12","12.12","12.12","12.12","12.12","","","","","","","12.12"],
			["",""," $10.0  Total","","","","","","","","12.12","12.12","12.12","12.12","12.12","12.12","","","","","","","12.12"],
			["","02/04/05 Total","","","","","","","","","12.12","12.12","12.12","12.12","12.12","12.12","","","","","","","12.12"],
			["Golf Total","","","","","","","","","","12.12","12.12","12.12","12.12","12.12","12.12","","","","","","","12.12"],
			["Tee","02/03/05","$ 10.0","1.0E+01","11.27","11.27","11.27","11.27","11.27","11.27","","","","","","","","","","","","","11.27"],
			["",""," $10.0  Total","","11.27","11.27","11.27","11.27","11.27","11.27","","","","","","","","","","","","","11.27"],
			["","02/03/05 Total","","","11.27","11.27","11.27","11.27","11.27","11.27","","","","","","","","","","","","","11.27"],
			["Tee Total","","","","11.27","11.27","11.27","11.27","11.27","11.27","","","","","","","","","","","","","11.27"],
			["Grand Total","","","","11.27","11.27","11.27","11.27","11.27","11.27","12.12","12.12","12.12","12.12","12.12","12.12","13.74","13.74","13.74","13.74","13.74","13.74","37.13"]
		]
	};

	function setPivotLayout(pivot, layout) {
		var props = new Asc.CT_pivotTableDefinition();
		props.ascHideValuesRow = true;
		switch (layout) {
			case "compact":
				props.asc_setCompact(true);
				props.asc_setOutline(true);
				break;
			case "outline":
				props.asc_setCompact(false);
				props.asc_setOutline(true);
				break;
			case "tabular":
				props.asc_setCompact(false);
				props.asc_setOutline(false);
				break;
			case "gridDropZones":
				props.asc_setCompact(false);
				props.asc_setOutline(false);
				props.asc_setGridDropZones(true);
				break;
			case "showHeaders":
				props.asc_setCompact(true);
				props.asc_setOutline(true);
				props.asc_setShowHeaders(false);
				break;

		}
		pivot.asc_set(api, props);
	}

	function testValidations() {
		QUnit.test("Test: Validations ", function(assert ) {
			assert.strictEqual(Asc.CT_pivotTableDefinition.prototype.isValidDataRef(dataRef), true, "Validations");
		});
	}

	function testCreate(prefix, init) {
		QUnit.test("Test: Layout " + prefix, function(assert ) {
			var pivot = init();
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);

			AscCommon.History.Clear();
			checkReportValues(assert, pivot, standards[prefix + "_0data"], "0data");

			pivot = checkHistoryOperation(assert, pivot, standards[prefix + "_1data"], "1data", function(){
				pivot.asc_addDataField(api, 5);
			});

			pivot = checkHistoryOperation(assert, pivot, standards[prefix + "_2data_col"], "2data_col", function(){
				pivot.asc_addDataField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards[prefix + "_2data_row"], "2data_row", function(){
				pivot.asc_moveToRowField(api, Asc.st_VALUES);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testLayout(layout) {
		testCreate(layout + "_0row_0col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			return pivot;
		});

		testCreate(layout + "_0row_1col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			pivot.asc_addColField(api, 2);
			return pivot;
		});

		testCreate(layout + "_0row_2col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			pivot.asc_addColField(api, 2);
			pivot.asc_addColField(api, 4);
			return pivot;
		});

		testCreate(layout + "_1row_0col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			pivot.asc_addRowField(api, 0);
			return pivot;
		});

		testCreate(layout + "_1row_1col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addColField(api, 2);
			return pivot;
		});

		testCreate(layout + "_1row_2col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addColField(api, 2);
			pivot.asc_addColField(api, 4);
			return pivot;
		});

		testCreate(layout + "_2row_0col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			return pivot;
		});

		testCreate(layout + "_2row_1col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addColField(api, 2);
			return pivot;
		});

		testCreate(layout + "_2row_2col", function() {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, layout);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addColField(api, 2);
			pivot.asc_addColField(api, 4);
			return pivot;
		});
	}

	function testLayoutValues(layout) {
		QUnit.test("Test: Layout Values " + layout, function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			setPivotLayout(pivot, layout);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addColField(api, 2);
			pivot.asc_addColField(api, 4);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addDataField(api, 6);

			AscCommon.History.Clear();
			checkReportValues(assert, pivot, standards[layout + "_2row_2col_2data_col"], "col3");

			pivot = checkHistoryOperation(assert, pivot, standards[layout + "_2row_2col_2data_col2"], "col2", function(){
				pivot.asc_moveColField(api, 2, 1);
			});

			pivot = checkHistoryOperation(assert, pivot, standards[layout + "_2row_2col_2data_col1"], "col1", function(){
				pivot.asc_moveColField(api, 1, 0);
			});

			pivot = checkHistoryOperation(assert, pivot, standards[layout + "_2row_2col_2data_row"], "row3", function(){
				pivot.asc_moveToRowField(api, Asc.st_VALUES);
			});

			pivot = checkHistoryOperation(assert, pivot, standards[layout + "_2row_2col_2data_row2"], "row2", function(){
				pivot.asc_moveRowField(api, 2, 1);
			});

			pivot = checkHistoryOperation(assert, pivot, standards[layout + "_2row_2col_2data_row1"], "row1", function(){
				pivot.asc_moveRowField(api, 1, 0);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testLayoutSubtotal(layout) {
		QUnit.test("Test: Subtotal " + layout, function(assert ) {
			var props;
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			setPivotLayout(pivot, layout);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addDataField(api, 6);
			pivot.asc_moveToRowField(api, Asc.st_VALUES);
			pivot.asc_addRowField(api, 2);
			pivot.asc_addRowField(api, 4);

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["subtotal_" + layout + "_none"], "none", function(){
				props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDefaultSubtotal(false);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["subtotal_" + layout + "_bottom"], "bottom", function(){
				props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDefaultSubtotal(true);
				props.asc_setSubtotalTop(false);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["subtotal_" + layout + "_top"], "top", function(){
				props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDefaultSubtotal(true);
				props.asc_setSubtotalTop(true);
				pivot.asc_set(api, props);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testPivotInsertBlankRow() {
		QUnit.test("Test: InsertBlankRow", function(assert) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.checkPivotFieldItems(0);
			pivot.checkPivotFieldItems(1);
			pivot.checkPivotFieldItems(2);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addDataField(api, 6);
			var props = new Asc.CT_pivotTableDefinition();
			props.ascHideValuesRow = true;
			props.asc_setInsertBlankRow(true);
			pivot.asc_set(api, props);

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["insertBlankRow_1row"], "1row", function(){
				pivot.asc_addRowField(api, 0);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["insertBlankRow_2row"], "2row", function(){
				pivot.asc_moveToRowField(api, Asc.st_VALUES);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["insertBlankRow_3row"], "3row", function(){
				pivot.asc_addRowField(api, 1);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["insertBlankRow_4row"], "4row", function(){
				pivot.asc_addRowField(api, 2);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testPivotPageFilterLayout() {
		QUnit.test("Test: PageFilter layout", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.pivotTableDefinitionX14 = new Asc.CT_pivotTableDefinitionX14();

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["filter_downThenOver1"], "downThenOver1", function(){
				pivot.asc_addPageField(api, 0);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["filter_downThenOver3"], "downThenOver3", function(){
				pivot.asc_addPageField(api, 1);
				pivot.asc_addPageField(api, 2);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["filter_downThenOver3_2wrap"], "downThenOver3_2wrap", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setPageWrap(2);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["filter_downThenOver7_2wrap"], "downThenOver7_2wrap", function(){
				pivot.asc_addPageField(api, 3);
				pivot.asc_addPageField(api, 4);
				pivot.asc_addPageField(api, 5);
				pivot.asc_addPageField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["filter_overThenDown7_2wrap"], "overThenDown7_2wrap", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setPageOverThenDown(true);
				pivot.asc_set(api, props);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testFieldProperty() {
		QUnit.test("Test: Field Property", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 4);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addDataField(api, 6);
			pivot.asc_moveToRowField(api, Asc.st_VALUES);
			pivot.asc_addRowField(api, 2);

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["field0"], "field0", function() {
				var pivotField = pivot.asc_getPivotFields()[0];
				var props = new Asc.CT_PivotField();
				props.asc_setCompact(false);
				props.asc_setOutline(true);
				props.asc_setSubtotalTop(false);
				props.asc_setInsertBlankRow(true);
				pivotField.asc_set(api, pivot, 0, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["field4"], "field4", function() {
				var pivotField = pivot.asc_getPivotFields()[4];
				var props = new Asc.CT_PivotField();
				props.asc_setName("UnitsCustom");
				props.asc_setOutline(false);
				props.asc_setCompact(false);
				props.asc_setShowAll(true);
				pivotField.asc_set(api, pivot, 4, props);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testFieldSubtotal() {
		QUnit.test("Test: Field Subtotal", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addColField(api, 2);
			pivot.asc_addDataField(api, 5);

			AscCommon.History.Clear();

			pivot = checkHistoryOperation(assert, pivot, standards["fieldSubtotalNone"], "fieldSubtotalNone", function() {
				var pivotField = pivot.asc_getPivotFields()[0];
				var props = new Asc.CT_PivotField();
				props.asc_setDefaultSubtotal(false);
				pivotField.asc_set(api, pivot, 0, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["fieldSubtotalAuto"], "fieldSubtotalAuto", function() {
				var pivotField = pivot.asc_getPivotFields()[0];
				var props = new Asc.CT_PivotField();
				props.asc_setDefaultSubtotal(true);
				pivotField.asc_set(api, pivot, 0, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["fieldSubtotalCustom"], "fieldSubtotalCustom", function() {
				var pivotField = pivot.asc_getPivotFields()[0];
				var props = new Asc.CT_PivotField();
				props.asc_setDefaultSubtotal(true);
				props.asc_setSubtotals([Asc.c_oAscItemType.Count]);
				pivotField.asc_set(api, pivot, 0, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["fieldSubtotalAll"], "fieldSubtotalAll", function() {
				var pivotField = pivot.asc_getPivotFields()[0];
				var props = new Asc.CT_PivotField();
				props.asc_setDefaultSubtotal(true);
				props.asc_setSubtotals([Asc.c_oAscItemType.Sum, Asc.c_oAscItemType.Count, Asc.c_oAscItemType.Avg,
					Asc.c_oAscItemType.Max, Asc.c_oAscItemType.Min, Asc.c_oAscItemType.Product,
					Asc.c_oAscItemType.CountA, Asc.c_oAscItemType.StdDev, Asc.c_oAscItemType.StdDevP,
					Asc.c_oAscItemType.Var, Asc.c_oAscItemType.VarP]);
				pivotField.asc_set(api, pivot, 0, props);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testDataValues() {
		QUnit.test("Test: data values", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.setHideValuesRow(true);

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["data_values1"], "values1", function() {
				var i;
				var types = [Asc.c_oAscDataConsolidateFunction.Count, Asc.c_oAscDataConsolidateFunction.CountNums, Asc.c_oAscDataConsolidateFunction.Min,
					Asc.c_oAscDataConsolidateFunction.Max, Asc.c_oAscDataConsolidateFunction.Sum, Asc.c_oAscDataConsolidateFunction.Average,
					Asc.c_oAscDataConsolidateFunction.Product, Asc.c_oAscDataConsolidateFunction.StdDev, Asc.c_oAscDataConsolidateFunction.StdDevp,
					Asc.c_oAscDataConsolidateFunction.Var, Asc.c_oAscDataConsolidateFunction.Varp]
				for (i = 0; i < types.length; ++i) {
					pivot.asc_addDataField(api, 5);
				}
				var dataFields = pivot.asc_getDataFields();
				for (i = 0; i < types.length; ++i) {
					var dataField = dataFields[i];
					var props = new Asc.CT_DataField();
					props.asc_setName(AscCommonExcel.ToName_ST_DataConsolidateFunction(types[i]) + " of " + pivot.getPivotFieldName(5) + (i + 1));
					props.asc_setSubtotal(types[i]);
					dataField.asc_set(api, pivot, i, props);
				}
			});

			pivot = checkHistoryOperation(assert, pivot, standards["data_values2"], "values2", function() {
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRef2);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["data_values3"], "values3", function() {
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRef3);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["data_values4"], "values4", function() {
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRef4);
				pivot.asc_set(api, props);
			});

			// TODO Fix dataRow count logic
			pivot = checkHistoryOperation(assert, pivot, standards["data_values5"], "values5", function() {
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRef5);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["data_values6"], "values6", function() {
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRef6);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["data_values7"], "values7", function() {
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRef7);
				pivot.asc_set(api, props);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testHeaderRename() {
		QUnit.test("Test: header rename", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRefHeader, ws, reportRange);
			var props = new Asc.CT_pivotTableDefinition();
			props.ascHideValuesRow = true;
			pivot.asc_set(api, props);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.checkPivotFieldItems(0);
			pivot.checkPivotFieldItems(1);
			pivot.checkPivotFieldItems(2);
			pivot.checkPivotFieldItems(3);
			pivot.checkPivotFieldItems(4);
			pivot.checkPivotFieldItems(5);
			pivot.checkPivotFieldItems(6);
			pivot.checkPivotFieldItems(7);
			pivot.checkPivotFieldItems(8);
			pivot.checkPivotFieldItems(9);
			pivot.checkPivotFieldItems(10);
			pivot.checkPivotFieldItems(11);
			pivot.checkPivotFieldItems(12);
			pivot.checkPivotFieldItems(13);

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["headerRename"], "header rename", function(){
				pivot.asc_addPageField(api, 0);
				pivot.asc_addPageField(api, 1);
				pivot.asc_addPageField(api, 2);
				pivot.asc_addPageField(api, 3);
				pivot.asc_addPageField(api, 4);
				pivot.asc_addPageField(api, 5);
				pivot.asc_addPageField(api, 6);
				pivot.asc_addPageField(api, 7);
				pivot.asc_addPageField(api, 8);
				pivot.asc_addPageField(api, 9);
				pivot.asc_addPageField(api, 10);
				pivot.asc_addPageField(api, 11);
				pivot.asc_addPageField(api, 12);
				pivot.asc_addPageField(api, 13);
				pivot.asc_addDataField(api, 0);
				pivot.asc_addDataField(api, 0);
				pivot.asc_addDataField(api, 1);
				pivot.asc_addDataField(api, 1);
			});
			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testPivotManipulationField() {
		QUnit.test("Test: Field Manipulation", function(assert) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			var props = new Asc.CT_pivotTableDefinition();
			props.ascHideValuesRow = true;
			pivot.asc_set(api, props);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.checkPivotFieldItems(0);
			pivot.checkPivotFieldItems(1);
			pivot.checkPivotFieldItems(2);
			pivot.checkPivotFieldItems(3);
			pivot.checkPivotFieldItems(4);
			pivot.checkPivotFieldItems(5);
			pivot.checkPivotFieldItems(6);

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["addField"], "addField", function(){
				pivot.asc_addField(api, 0);
				pivot.asc_addField(api, 1);
				pivot.asc_addField(api, 2);
				pivot.asc_addField(api, 3);
				pivot.asc_addField(api, 4);
				pivot.asc_addField(api, 5);
				pivot.asc_addField(api, 6);
				pivot.asc_addDataField(api, 0);
				pivot.asc_addDataField(api, 1);
				pivot.asc_addDataField(api, 2);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["moveDataField"], "moveDataField", function(){
				pivot.asc_addDataField(api, 5);
				pivot.asc_moveToPageField(api, 5, 2);
				pivot.asc_moveToColField(api, 0, 3);
				pivot.asc_moveToRowField(api, 3, 0);
				pivot.asc_moveToDataField(api, 4, 0);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["moveToField"], "moveToField", function(){
				pivot.asc_moveToPageField(api, 2);
				pivot.asc_moveToColField(api, 5);
				pivot.asc_moveToDataField(api, 0);
				pivot.asc_moveToRowField(api, 4, 5);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["compact_0row_0col_0data"], "removeField", function(){
				pivot.asc_removeField(api, 0);
				pivot.asc_removeField(api, 1);
				pivot.asc_removeField(api, 2);
				pivot.asc_removeField(api, 3);
				pivot.asc_removeField(api, 4);
				pivot.asc_removeField(api, 5);
				pivot.asc_removeField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["moveField"], "moveField", function(){
				pivot.asc_addRowField(api, 0);
				pivot.asc_addRowField(api, 1);
				pivot.asc_addPageField(api, 2);
				pivot.asc_addColField(api, 3);
				pivot.asc_addColField(api, 4);
				pivot.asc_addPageField(api, 5);
				pivot.asc_addPageField(api, 6);
				pivot.asc_addDataField(api, 5);
				pivot.asc_addDataField(api, 6);
				pivot.asc_addDataField(api, 0);
				pivot.asc_addDataField(api, 1);
				pivot.asc_addDataField(api, 5);

				pivot.asc_movePageField(api, 0, 1);
				pivot.asc_moveRowField(api, 1, 0);
				pivot.asc_moveColField(api, 0, 2);
				pivot.asc_moveDataField(api, 1, 3);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["removeDataField"], "removeDataField", function(){
				pivot.asc_removeNoDataField(api, 0);
				pivot.asc_removeDataField(api, 5, 0);
				pivot.asc_removeField(api, 6);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testPivotManipulationValues() {
		QUnit.test("Test: manipulation values", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			var props = new Asc.CT_pivotTableDefinition();
			props.ascHideValuesRow = true;
			pivot.asc_set(api, props);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addRowField(api, 2);
			pivot.asc_addDataField(api, 5);


			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["dataOnCols"], "dataOnCols1", function(){
				pivot.asc_addDataField(api, 6);
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addDataField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["dataOnRows"], "dataOnRows2", function(){
				pivot.asc_moveToPageField(api, Asc.st_VALUES);
				pivot.asc_moveToColField(api, Asc.st_VALUES);
				pivot.asc_moveToDataField(api, Asc.st_VALUES);
				pivot.asc_moveToRowField(api, Asc.st_VALUES);
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addDataField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["dataPosition"], "dataPosition3", function(){
				pivot.asc_moveRowField(api, 3, 2);
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addDataField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["dataPosition"], "dataPosition22", function(){
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addDataField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["dataPosition"], "dataPosition4", function(){
				pivot.asc_removeField(api, 0);
				pivot.asc_removeField(api, 1);
				pivot.asc_removeField(api, 2);
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addRowField(api, 0);
				pivot.asc_addRowField(api, 1);
				pivot.asc_addRowField(api, 2);
				pivot.asc_addDataField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["dataOnRows"], "dataOnRows5", function(){
				pivot.asc_moveToColField(api, Asc.st_VALUES);
				pivot.asc_moveToRowField(api, Asc.st_VALUES);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["dataOnRows"], "dataOnRows6", function(){
				pivot.asc_moveRowField(api, 3, 2);
				pivot.asc_removeField(api, 2);
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addDataField(api, 6);
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addRowField(api, 2);
				pivot.asc_addDataField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["dataOnRows"], "dataOnRows6", function(){
				pivot.asc_moveRowField(api, 3, 2);
				pivot.asc_removeField(api, 2);
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addDataField(api, 6);
				pivot.asc_removeDataField(api, 6, 1);
				pivot.asc_addRowField(api, 2);
				pivot.asc_addDataField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["dataOnRowsValues"], "dataOnRowsValues", function(){
				pivot.asc_removeNoDataField(api, Asc.st_VALUES);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testDataRefresh() {
		QUnit.test("Test: data refresh", function(assert ) {
			var pivotField, props;
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addColField(api, 2);
			pivot.asc_addColField(api, 3);
			pivot.asc_addPageField(api, 4);
			pivot.asc_addPageField(api, 6);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addDataField(api, 0);
			pivot.asc_addDataField(api, 6);

			pivotField = pivot.asc_getPivotFields()[0];
			props = new Asc.CT_PivotField();
			props.asc_setCompact(false);
			props.asc_setOutline(false);
			props.asc_setSubtotals([Asc.c_oAscItemType.Avg]);
			pivotField.asc_set(api, pivot, 0, props);

			pivotField = pivot.asc_getPivotFields()[2];
			props = new Asc.CT_PivotField();
			props.asc_setDefaultSubtotal(false);
			pivotField.asc_set(api, pivot, 2, props);

			pivotField = pivot.asc_getPivotFields()[4];
			props = new Asc.CT_PivotField();
			props.asc_setName("RenamedUnits");
			pivotField.asc_set(api, pivot, 4, props);

			var dataField = pivot.asc_getDataFields()[0];
			props = new Asc.CT_DataField();
			props.asc_setName(AscCommonExcel.ToName_ST_DataConsolidateFunction(Asc.c_oAscItemType.Avg) + " of " + pivot.getPivotFieldName(5));
			props.asc_setSubtotal(Asc.c_oAscItemType.Avg);
			dataField.asc_set(api, pivot, 0, props);


			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["refreshFieldSettings"], "refreshFieldSettings", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRefFieldSettings);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["refreshRecords"], "refreshRecords", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRefRecords);
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["refreshStructure"], "refreshStructure", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRefStructure);
				pivot.asc_set(api, props);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testDataSource() {
		QUnit.test("Test: data source", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRefTable, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.checkPivotFieldItems(0);
			pivot.checkPivotFieldItems(2);

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["compact_1row_1col_1data"], "table", function(){
				pivot.asc_addRowField(api, 0);
				pivot.asc_addColField(api, 2);
				pivot.asc_addDataField(api, 5);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["compact_0row_1col_1data"], "table columns", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRefTableColumn);
				pivot.asc_set(api, props);
				pivot.asc_removeField(api, 1);
				pivot.asc_removeField(api, 4);

				pivot.asc_addColField(api, 1);
				pivot.asc_addDataField(api, 4);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["compact_1row_1col_1data"], "def name", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRefDefName);
				pivot.asc_set(api, props);
				pivot.asc_removeField(api, 2);
				pivot.asc_removeField(api, 5);

				pivot.asc_addRowField(api, 0);
				pivot.asc_addColField(api, 2);
				pivot.asc_addDataField(api, 5);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["compact_1row_1col_1data"], "def name local", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setDataRef(dataRefDefNameLocal);
				pivot.asc_set(api, props);
				pivot.asc_removeField(api, 0);
				pivot.asc_removeField(api, 2);
				pivot.asc_removeField(api, 5);

				pivot.asc_addRowField(api, 0);
				pivot.asc_addColField(api, 2);
				pivot.asc_addDataField(api, 5);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testFiltersValueFilter() {
		QUnit.test("Test: filters value filter", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRefFilters, ws, reportRange);
			setPivotLayout(pivot, 'tabular');
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addRowField(api, 3);
			pivot.asc_addColField(api, 0);
			pivot.asc_addColField(api, 4);
			pivot.asc_addDataField(api, 5);

			var getNewFilter = function(val){
				var pivotFilterObj = new Asc.PivotFilterObj();
				pivotFilterObj.asc_setDataFieldIndexSorting(0);
				pivotFilterObj.asc_setDataFieldIndexFilter(1);
				pivotFilterObj.asc_setIsPageFilter(false);
				pivotFilterObj.asc_setIsMultipleItemSelectionAllowed(false);
				pivotFilterObj.asc_setIsTop10Sum(false);
				var filterObj = new Asc.AutoFilterObj();
				filterObj.asc_setFilter(new Asc.CustomFilters());
				filterObj.asc_setType(Asc.c_oAscAutoFilterTypes.CustomFilters);
				var customFilter = filterObj.asc_getFilter();
				customFilter.asc_setCustomFilters([new Asc.CustomFilter()]);
				customFilter.asc_setAnd(true);
				var customFilters = customFilter.asc_getCustomFilters();
				customFilters[0].asc_setOperator(Asc.c_oAscCustomAutoFilter.isGreaterThan);
				customFilters[0].asc_setVal(val);
				var autoFilterObject = new Asc.AutoFiltersOptions();
				autoFilterObject.pivotObj = pivotFilterObj;
				autoFilterObject.filter = filterObj;
				return autoFilterObject;
			};

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["valueFilterOrder1"], "order1", function(){
				pivot.filterByFieldIndex(api, getNewFilter("1"), 4, true);
				pivot.filterByFieldIndex(api, getNewFilter("2"), 3, true);
				pivot.filterByFieldIndex(api, getNewFilter("18"), 0, true);
				pivot.filterByFieldIndex(api, getNewFilter("20"), 1, true);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["valueFilterOrder2"], "order2", function(){
				pivot.asc_removeFilters(api);

				pivot.filterByFieldIndex(api, getNewFilter("20"), 1, true);
				pivot.filterByFieldIndex(api, getNewFilter("18"), 0, true);
				pivot.filterByFieldIndex(api, getNewFilter("1"), 4, true);
				pivot.filterByFieldIndex(api, getNewFilter("2"), 3, true);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testFiltersValueFilterBug46141() {
		QUnit.test("Test: value filter bug 46141", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, 'tabular');
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addRowField(api, 2);
			pivot.asc_addDataField(api, 5);

			var getNewFilter = function(val){
				var pivotFilterObj = new Asc.PivotFilterObj();
				pivotFilterObj.asc_setDataFieldIndexSorting(0);
				pivotFilterObj.asc_setDataFieldIndexFilter(1);
				pivotFilterObj.asc_setIsPageFilter(false);
				pivotFilterObj.asc_setIsMultipleItemSelectionAllowed(false);
				pivotFilterObj.asc_setIsTop10Sum(false);
				var filterObj = new Asc.AutoFilterObj();
				filterObj.asc_setFilter(new Asc.CustomFilters());
				filterObj.asc_setType(Asc.c_oAscAutoFilterTypes.CustomFilters);
				var customFilter = filterObj.asc_getFilter();
				customFilter.asc_setCustomFilters([new Asc.CustomFilter()]);
				customFilter.asc_setAnd(true);
				var customFilters = customFilter.asc_getCustomFilters();
				customFilters[0].asc_setOperator(Asc.c_oAscCustomAutoFilter.isGreaterThan);
				customFilters[0].asc_setVal(val);
				var autoFilterObject = new Asc.AutoFiltersOptions();
				autoFilterObject.pivotObj = pivotFilterObj;
				autoFilterObject.filter = filterObj;
				return autoFilterObject;
			};

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["bug-46141-row"], "rows", function(){
				pivot.filterByFieldIndex(api, getNewFilter("13.5"), 2, true);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["bug-46141-col"], "cols", function(){
				pivot.asc_moveToColField(api, 0);
				pivot.asc_moveToColField(api, 1);
				pivot.asc_moveToColField(api, 2);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}


	function testFiltersTop10() {
		QUnit.test("Test: filters top10", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, 'tabular');
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 2);
			pivot.asc_addRowField(api, 4);
			pivot.asc_addColField(api, 1);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addDataField(api, 6);

			var getNewFilter = function(top, isPercent, val, isTop10Sum, dataIndex){
				var pivotFilterObj = new Asc.PivotFilterObj();
				pivotFilterObj.asc_setDataFieldIndexSorting(0);
				pivotFilterObj.asc_setDataFieldIndexFilter(dataIndex);
				pivotFilterObj.asc_setIsPageFilter(false);
				pivotFilterObj.asc_setIsMultipleItemSelectionAllowed(false);
				pivotFilterObj.asc_setIsTop10Sum(isTop10Sum);
				var filterObj = new Asc.AutoFilterObj();
				filterObj.asc_setFilter(new Asc.Top10());
				filterObj.asc_setType(Asc.c_oAscAutoFilterTypes.Top10);
				var top10Filter = filterObj.asc_getFilter();
				top10Filter.asc_setTop(top);
				top10Filter.asc_setPercent(isPercent);
				top10Filter.asc_setVal(val);
				var autoFilterObject = new Asc.AutoFiltersOptions();
				autoFilterObject.pivotObj = pivotFilterObj;
				autoFilterObject.filter = filterObj;
				return autoFilterObject;
			};

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["top10"], "top10", function(){
				pivot.filterByFieldIndex(api, getNewFilter(true, false, 1, false, 1), 4, true);
				pivot.filterByFieldIndex(api, getNewFilter(false, true, 2, false, 2), 2, true);
				pivot.filterByFieldIndex(api, getNewFilter(true, false, 12, true, 1), 0, true);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testFiltersLabel() {
		QUnit.test("Test: filters label", function(assert) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, 'tabular');
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 4);
			pivot.asc_addRowField(api, 5);
			pivot.asc_addRowField(api, 6);
			pivot.asc_addDataField(api, 3);

			var getNewFilter = function(type1, val1, type2, val2, val3){
				var pivotFilterObj = new Asc.PivotFilterObj();
				pivotFilterObj.asc_setDataFieldIndexSorting(0);
				pivotFilterObj.asc_setDataFieldIndexFilter(val3);
				pivotFilterObj.asc_setIsPageFilter(false);
				pivotFilterObj.asc_setIsMultipleItemSelectionAllowed(false);
				pivotFilterObj.asc_setIsTop10Sum(false);
				var filterObj = new Asc.AutoFilterObj();
				filterObj.asc_setFilter(new Asc.CustomFilters());
				filterObj.asc_setType(Asc.c_oAscAutoFilterTypes.CustomFilters);
				var filter = filterObj.asc_getFilter();
				filter.asc_setAnd(true);
				var customFilters = [];
				var customFilter;
				customFilter= new Asc.CustomFilter();
				customFilter.asc_setOperator(type1);
				customFilter.asc_setVal(val1);
				customFilters.push(customFilter);
				if (undefined !== type2) {
					filter.asc_setAnd(false);
					customFilter = new Asc.CustomFilter();
					customFilter.asc_setOperator(type2);
					customFilter.asc_setVal(val2);
					customFilters.push(customFilter);
				}
				filter.asc_setCustomFilters(customFilters);
				var autoFilterObject = new  Asc.AutoFiltersOptions();
				autoFilterObject.pivotObj = pivotFilterObj;
				autoFilterObject.filter = filterObj;
				return autoFilterObject;
			};

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["label1"], "label1", function(){
				pivot.filterByFieldIndex(api, getNewFilter(Asc.c_oAscCustomAutoFilter.isGreaterThan, "10.6", undefined, undefined, 1), 6, true);
				pivot.filterByFieldIndex(api, getNewFilter(Asc.c_oAscCustomAutoFilter.contains, "3", undefined, undefined, 0), 5, true);
				//todo fix matchingValues("11", "11", "<>");
				// pivot.filterByFieldIndex(api, getNewFilter(Asc.c_oAscCustomAutoFilter.doesNotEqual, "11", undefined, undefined, 0), 4, true);
				pivot.filterByFieldIndex(api, getNewFilter(Asc.c_oAscCustomAutoFilter.isLessThan, "11", Asc.c_oAscCustomAutoFilter.isGreaterThan, "11", 0), 4, true);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testFiltersReIndex() {
		QUnit.test("Test: filters reIndex", function(assert) {
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			setPivotLayout(pivot, 'tabular');
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addColField(api, 0);
			pivot.asc_addColField(api, 1);
			pivot.asc_addRowField(api, 2);
			pivot.asc_addRowField(api, 4);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addDataField(api, 6);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addDataField(api, 6);

			var getNewFilter = function(val, iMeasureFld){
				var pivotFilterObj = new Asc.PivotFilterObj();
				pivotFilterObj.asc_setDataFieldIndexSorting(0);
				pivotFilterObj.asc_setDataFieldIndexFilter(iMeasureFld);
				pivotFilterObj.asc_setIsPageFilter(false);
				pivotFilterObj.asc_setIsMultipleItemSelectionAllowed(false);
				pivotFilterObj.asc_setIsTop10Sum(false);
				var filterObj = new Asc.AutoFilterObj();
				filterObj.asc_setFilter(new Asc.CustomFilters());
				filterObj.asc_setType(Asc.c_oAscAutoFilterTypes.CustomFilters);
				var customFilter = filterObj.asc_getFilter();
				customFilter.asc_setCustomFilters([new Asc.CustomFilter()]);
				customFilter.asc_setAnd(true);
				var customFilters = customFilter.asc_getCustomFilters();
				customFilters[0].asc_setOperator(Asc.c_oAscCustomAutoFilter.isGreaterThan);
				customFilters[0].asc_setVal(val);
				var autoFilterObject = new Asc.AutoFiltersOptions();
				autoFilterObject.pivotObj = pivotFilterObj;
				autoFilterObject.filter = filterObj;
				return autoFilterObject;
			};

			AscCommon.History.Clear();

			pivot = checkHistoryOperation(assert, pivot, standards["reIndexStart"], "reIndexStart", function(){
				pivot.filterByFieldIndex(api, getNewFilter("40", 1), 2, true);
				pivot.filterByFieldIndex(api, getNewFilter("46", 4), 0, true);
				pivot.sortByFieldIndex(api, 4, Asc.c_oAscSortOptions.Descending, 1);
				pivot.sortByFieldIndex(api, 1, Asc.c_oAscSortOptions.Ascending, 2);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["reIndexMove"], "reIndexMove", function(){
				pivot.asc_moveDataField(api, 3, 0);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["reIndexDelete"], "reIndexDelete", function(){
				pivot.asc_removeField(api, 6);
			});

			pivot = checkHistoryOperation(assert, pivot, standards["reIndexAdd"], "reIndexAdd", function(){
				pivot.asc_addDataField(api, 6, 1);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testNumFormat() {
		QUnit.test("Test: num format", function(assert ) {
			var pivot = api._asc_insertPivot(wb, dataRefNumFormat, ws, reportRange);
			setPivotLayout(pivot, 'tabular');
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 1);
			pivot.asc_addPageField(api, 2);
			pivot.asc_addRowField(api, 3);
			pivot.asc_addRowField(api, 4);
			pivot.asc_addColField(api, 5);
			pivot.asc_addDataField(api, 5);
			pivot.asc_addColField(api, 6);
			pivot.asc_addColField(api, 7);
			pivot.asc_addColField(api, 8);
			pivot.asc_addColField(api, 9);
			pivot.asc_addColField(api, 10);

			var getNewFilter = function(fld, index) {
				var autoFilterObject = new Asc.AutoFiltersOptions();
				pivot.fillAutoFiltersOptions(autoFilterObject, fld);
				for (var i = 0; i < autoFilterObject.values.length; ++i) {
					autoFilterObject.values[i].visible = i == index;
				}
				autoFilterObject.filter.type = Asc.c_oAscAutoFilterTypes.Filters;
				return autoFilterObject;
			};

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards["numFormat"], "numFormat", function(){
				pivot.filterByFieldIndex(api, getNewFilter(2, 0), 2, true);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}

	function testPivotMisc() {
		QUnit.test("Test: misc", function(assert ) {
			let testData =  [
				["Region","Gender","Style","Ship date","Units","Price","Cost"],
				["East","Boy","Tee","38383","12","11.04","10.42"]
			];
			let standards_compact_0row_0col_0data = standards["compact_0row_0col_0data"];
			let standards_longHeader = [
				["Sum of Price","Gender","Style"],
				["","Boy",""],
				["Region","Tee",""],
				["East","11.04",""]
			];
			let testDataRange = new Asc.Range(0, 0, testData[0].length - 1, testData.length - 1);
			fillData(wsData, testData, testDataRange);
			let dataRef = wsData.getName() + "!" + testDataRange.getName();

			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			var props = new Asc.CT_pivotTableDefinition();
			props.ascHideValuesRow = true;
			pivot.asc_set(api, props);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.pivotTableDefinitionX14 = new Asc.CT_pivotTableDefinitionX14();
			pivot.checkPivotFieldItems(0);
			pivot.checkPivotFieldItems(1);
			pivot.checkPivotFieldItems(2);

			AscCommon.History.Clear();
			pivot = checkHistoryOperation(assert, pivot, standards_compact_0row_0col_0data, "misc", function(){
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setName("new<&>pivot name");
				props.asc_setTitle("Title");
				props.asc_setDescription("Description");
				pivot.asc_set(api, props);
			});

			pivot = checkHistoryOperation(assert, pivot, standards_longHeader, "longHeader", function(){
				pivot.asc_addRowField(api, 0);
				pivot.asc_addColField(api, 1);
				pivot.asc_addColField(api, 2);
				pivot.asc_addDataField(api, 5);
				var props = new Asc.CT_pivotTableDefinition();
				props.ascHideValuesRow = true;
				props.asc_setCompact(false);
				props.asc_setOutline(true);
				props.asc_setRowGrandTotals(false);
				props.asc_setColGrandTotals(false);
				pivot.asc_set(api, props);
				var pivotField = pivot.asc_getPivotFields()[1];
				props = new Asc.CT_PivotField();
				props.asc_setDefaultSubtotal(false);
				pivotField.asc_set(api, pivot, 1, props);
			});

			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());

			// prot["asc_setShowHeaders"] = prot.asc_setShowHeaders;
			// prot["asc_setFillDownLabelsDefault"] = prot.asc_setFillDownLabelsDefault;
		});
	}
	function testPivotShowAs() {
		QUnit.test("Test: Show as", function(assert) {
			let testData =  [
				["Region","Gender","Style","Ship date","Units","Price","Cost"],
				["East","Boy","Tee","38383","12","11.04","10.42"],
				["East","Boy","Golf","38383","12","13","12.6"],
				["East","Boy","Fancy","38383","12","11.96","11.74"],
				["East","Girl","Tee","38383","10","11.27","10.56"],
				["East","Girl","Golf","38383","10","12.12","11.95"],
				["East","Girl","Fancy","38383","10","13.74","13.33"],
				["North","Boy","Tee","38383","16","13.08","14.06"],
				["North","Helicopter","Tee","38383","16","5555","14.06"],
				["West","Boy","Tee","38383","11","11.44","10.94"],
				["West","Boy","Golf","38383","11","12.63","11.73"],
				["West","Boy","Fancy","38383","11","12.06","11.51"],
				["West","Girl","Tee","38383","15","13.42","13.29"],
				["West","Girl","Golf","38383","15","11.48","10.67"]
			];
			let testDataRange = new Asc.Range(0, 0, testData[0].length - 1, testData.length - 1);
			fillData(wsData, testData, testDataRange);
			let dataRef = wsData.getName() + "!" + testDataRange.getName();
			let percentOfTotal_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels", "Boy","Girl","Helicopter","Grand Total"],
				["East","0.006313308","0.006511476","0","0.012824785"],
				["Fancy","0.002097421","0.002409579","0","0.004507001"],
				["Golf","0.002279806","0.002125481","0","0.004405286"],
				["Tee","0.001936081","0.001976416","0","0.003912498"],
				["North","0.002293835","0","0.974178568","0.976472404"],
				["Tee","0.002293835","0","0.974178568","0.976472404"],
				["West","0.006336107","0.004366705","0","0.010702812"],
				["Fancy","0.002114958","0","0","0.002114958"],
				["Golf","0.002214919","0.002013244","0","0.004228163"],
				["Tee","0.002006229","0.002353461","0","0.00435969"],
				["Grand Total","0.01494325","0.010878181","0.974178568","1"]
			  ];
			let differenceNext_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","22.92","37.13","-5555","-5494.95"],
				["Fancy","-0.1","13.74","0","13.64"],
				["Golf","0.37","0.64","0","1.01"],
				["Tee","-2.04","11.27","-5555","-5545.77"],
				["North","-23.05","-24.9","5555","5507.05"],
				["Tee","1.64","-13.42","5555","5543.22"],
				["West","","","",""],
				["Fancy","","","",""],
				["Golf","","","",""],
				["Tee","","","",""],
				["Grand Total","","","",""]
				];
			let differenceNext_tabular = [
				["Sum of Price","","Gender","","",""],
				["Region","Style","Boy","Girl","Helicopter","Grand Total"],
				["East","Fancy","-0.1","13.74","0","13.64"],
				["","Golf","0.37","0.64","0","1.01"],
				["","Tee","-2.04","11.27","-5555","-5545.77"],
				["East Total","","22.92","37.13","-5555","-5494.95"],
				["North","Tee","1.64","-13.42","5555","5543.22"],
				["North Total","","-23.05","-24.9","5555","5507.05"],
				["West","Fancy","","","",""],
				["","Golf","","","",""],
				["","Tee","","","",""],
				["West Total","","","","",""],
				["Grand Total","","","","",""]
				];
			let differenceNext_compact2 = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","-1.13","37.13","",""],
				["Fancy","-1.78","13.74","",""],
				["Golf","0.88","12.12","",""],
				["Tee","-0.23","11.27","",""],
				["North","13.08","-5555","",""],
				["Tee","13.08","-5555","",""],
				["West","11.23","24.9","",""],
				["Fancy","12.06","0","",""],
				["Golf","1.15","11.48","",""],
				["Tee","-1.98","13.42","",""],
				["Grand Total","23.18","-5492.97","",""]
				];
			let differencePrev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","","","",""],
				["Fancy","","","",""],
				["Golf","","","",""],
				["Tee","","","",""],
				["North","-22.92","-37.13","5555","5494.95"],
				["Tee","2.04","-11.27","5555","5545.77"],
				["West","23.05","24.9","-5555","-5507.05"],
				["Fancy","0.1","-13.74","0","-13.64"],
				["Golf","-0.37","-0.64","0","-1.01"],
				["Tee","-1.64","13.42","-5555","-5543.22"],
				["Grand Total","","","",""]
				];
			let differencePrev_compact2 = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","","1.13","-37.13",""],
				["Fancy","","1.78","-13.74",""],
				["Golf","","-0.88","-12.12",""],
				["Tee","","0.23","-11.27",""],
				["North","","-13.08","5555",""],
				["Tee","","-13.08","5555",""],
				["West","","-11.23","-24.9",""],
				["Fancy","","-12.06","0",""],
				["Golf","","-1.15","-11.48",""],
				["Tee","","1.98","-13.42",""],
				["Grand Total","","-23.18","5492.97",""]
				];
			let differenceBase_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","22.92","37.13","-5555","-5494.95"],
				["Fancy","#N/A","#N/A","#N/A","#N/A"],
				["Golf","#N/A","#N/A","#N/A","#N/A"],
				["Tee","-2.04","11.27","-5555","-5545.77"],
				["North","","","",""],
				["Tee","","","",""],
				["West","23.05","24.9","-5555","-5507.05"],
				["Fancy","#N/A","#N/A","#N/A","#N/A"],
				["Golf","#N/A","#N/A","#N/A","#N/A"],
				["Tee","-1.64","13.42","-5555","-5543.22"],
				["Grand Total","","","",""]
				];
			let differenceBase_compact2 = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","-1.13","","-37.13",""],
				["Fancy","-1.78","","-13.74",""],
				["Golf","0.88","","-12.12",""],
				["Tee","-0.23","","-11.27",""],
				["North","13.08","","5555",""],
				["Tee","13.08","","5555",""],
				["West","11.23","","-24.9",""],
				["Fancy","12.06","","0",""],
				["Golf","1.15","","-11.48",""],
				["Tee","-1.98","","-13.42",""],
				["Grand Total","23.18","","5492.97",""]
				]
			let percentOfCol_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","0.422485624","0.598581332","0","0.012824785"],
				["Fancy","0.140359113","0.221505723","0","0.004507001"],
				["Golf","0.152564253","0.195389328","0","0.004405286"],
				["Tee","0.129562258","0.181686281","0","0.003912498"],
				["North","0.15350311","0","1","0.976472404"],
				["Tee","0.15350311","0","1","0.976472404"],
				["West","0.424011266","0.401418668","0","0.010702812"],
				["Fancy","0.141532684","0","0","0.002114958"],
				["Golf","0.14822204","0.185071739","0","0.004228163"],
				["Tee","0.134256543","0.216346929","0","0.00435969"],
				["Grand Total","1","1","1","1"]
				];
			let percentOfRow_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","0.492274033","0.507725967","0","1"],
				["Fancy","0.46536965","0.53463035","0","1"],
				["Golf","0.517515924","0.482484076","0","1"],
				["Tee","0.494845361","0.505154639","0","1"],
				["North","0.002349104","0","0.997650896","1"],
				["Tee","0.002349104","0","0.997650896","1"],
				["West","0.592003932","0.407996068","0","1"],
				["Fancy","1","0","0","1"],
				["Golf","0.523849025","0.476150975","0","1"],
				["Tee","0.460176991","0.539823009","0","1"],
				["Grand Total","0.01494325","0.010878181","0.974178568","1"]
				];
			let percentOfParentRow_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","0.422485624","0.598581332","0","0.012824785"],
				["Fancy","0.332222222","0.370051172","","0.351428962"],
				["Golf","0.361111111","0.326420684","","0.34349788"],
				["Tee","0.306666667","0.303528144","","0.305073157"],
				["North","0.15350311","0","1","0.976472404"],
				["Tee","1","","1","1"],
				["West","0.424011266","0.401418668","0","0.010702812"],
				["Fancy","0.333794631","0","","0.197607734"],
				["Golf","0.349570994","0.461044177","","0.395051614"],
				["Tee","0.316634376","0.538955823","","0.407340652"],
				["Grand Total","1","1","1","1"]
				];
			let index_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","32.942902","46.67379205","0","1"],
				["Fancy","31.14246487","49.14703479","0","1"],
				["Golf","34.63208544","44.3533774","0","1"],
				["Tee","33.11497489","46.43741721","0","1"],
				["North","0.157201688","0","1.024094481","1"],
				["Tee","0.157201688","0","1.024094481","1"],
				["West","39.61681145","37.50590837","0","1"],
				["Fancy","66.91984509","0","0","1"],
				["Golf","35.05589562","43.77119352","0","1"],
				["Tee","30.79497296","49.62438101","0","1"],
				["Grand Total","1","1","1","1"]
				];
			let percentOfParentCol_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","0.492274033","0.507725967","0","1"],
				["Fancy","0.46536965","0.53463035","0","1"],
				["Golf","0.517515924","0.482484076","0","1"],
				["Tee","0.494845361","0.505154639","0","1"],
				["North","0.002349104","0","0.997650896","1"],
				["Tee","0.002349104","0","0.997650896","1"],
				["West","0.592003932","0.407996068","0","1"],
				["Fancy","1","0","0","1"],
				["Golf","0.523849025","0.476150975","0","1"],
				["Tee","0.460176991","0.539823009","0","1"],
				["Grand Total","0.01494325","0.010878181","0.974178568","1"]
				];
			let percentOfParent_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","1","1","","1"],
				["Fancy","0.332222222","0.370051172","","0.351428962"],
				["Golf","0.361111111","0.326420684","","0.34349788"],
				["Tee","0.306666667","0.303528144","","0.305073157"],
				["North","1","","1","1"],
				["Tee","1","","1","1"],
				["West","1","1","","1"],
				["Fancy","0.333794631","0","","0.197607734"],
				["Golf","0.349570994","0.461044177","","0.395051614"],
				["Tee","0.316634376","0.538955823","","0.407340652"],
				["Grand Total","","","",""]
				];
			let percentNext_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","2.752293578","","#NULL!","0.013133791"],
				["Fancy","0.991708126","","#NULL!","2.131011609"],
				["Golf","1.029295329","1.055749129","#NULL!","1.041891331"],
				["Tee","0.844036697","","#NULL!","0.004006767"],
				["North","0.362026017","#NULL!","","91.23513026"],
				["Tee","1.143356643","#NULL!","","223.9774739"],
				["West","1","1","","1"],
				["Fancy","1","","","1"],
				["Golf","1","1","","1"],
				["Tee","1","1","","1"],
				["Grand Total","","","",""]
				];
			let percentPrev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","1","1","","1"],
				["Fancy","1","1","","1"],
				["Golf","1","1","","1"],
				["Tee","1","1","","1"],
				["North","0.363333333","#NULL!","","76.13947764"],
				["Tee","1.184782609","#NULL!","","249.5777678"],
				["West","2.762232416","","#NULL!","0.01096069"],
				["Fancy","1.008361204","#NULL!","#NULL!","0.4692607"],
				["Golf","0.971538462","0.947194719","#NULL!","0.959792994"],
				["Tee","0.874617737","","#NULL!","0.004464735"],
				["Grand Total","","","",""]
				];
			let percentBase_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","2.752293578","","#NULL!","0.013133791"],
				["Fancy","#N/A","#N/A","#N/A","#N/A"],
				["Golf","#N/A","#N/A","#N/A","#N/A"],
				["Tee","0.844036697","","#NULL!","0.004006767"],
				["North","1","","1","1"],
				["Tee","1","","1","1"],
				["West","2.762232416","","#NULL!","0.01096069"],
				["Fancy","#N/A","#N/A","#N/A","#N/A"],
				["Golf","#N/A","#N/A","#N/A","#N/A"],
				["Tee","0.874617737","","#NULL!","0.004464735"],
				["Grand Total","","","",""]
				];
			let percentDiffNext_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","1.752293578","","#NULL!","-0.986866209"],
				["Fancy","-0.008291874","","#NULL!","1.131011609"],
				["Golf","0.029295329","0.055749129","#NULL!","0.041891331"],
				["Tee","-0.155963303","","#NULL!","-0.995993233"],
				["North","-0.637973983","#NULL!","","90.23513026"],
				["Tee","0.143356643","#NULL!","","222.9774739"],
				["West","","","",""],
				["Fancy","","","",""],
				["Golf","","","",""],
				["Tee","","","",""],
				["Grand Total","","","",""]
				];
			let percentDiffPrev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","","","",""],
				["Fancy","","","",""],
				["Golf","","","",""],
				["Tee","","","",""],
				["North","-0.636666667","#NULL!","","75.13947764"],
				["Tee","0.184782609","#NULL!","","248.5777678"],
				["West","1.762232416","","#NULL!","-0.98903931"],
				["Fancy","0.008361204","#NULL!","#NULL!","-0.5307393"],
				["Golf","-0.028461538","-0.052805281","#NULL!","-0.040207006"],
				["Tee","-0.125382263","","#NULL!","-0.995535265"],
				["Grand Total","","","",""]
				];
			let percentDiffBase_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","1.752293578","","#NULL!","-0.986866209"],
				["Fancy","#N/A","#N/A","#N/A","#N/A"],
				["Golf","#N/A","#N/A","#N/A","#N/A"],
				["Tee","-0.155963303","","#NULL!","-0.995993233"],
				["North","","","",""],
				["Tee","","","",""],
				["West","1.762232416","","#NULL!","-0.98903931"],
				["Fancy","#N/A","#N/A","#N/A","#N/A"],
				["Golf","#N/A","#N/A","#N/A","#N/A"],
				["Tee","-0.125382263","","#NULL!","-0.995535265"],
				["Grand Total","","","",""]
				];
			let runTotal_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","36","37.13","0","73.13"],
				["Fancy","11.96","13.74","0","25.7"],
				["Golf","13","12.12","0","25.12"],
				["Tee","11.04","11.27","0","22.31"],
				["North","49.08","37.13","5555","5641.21"],
				["Tee","24.12","11.27","5555","5590.39"],
				["West","85.21","62.03","5555","5702.24"],
				["Fancy","24.02","13.74","0","37.76"],
				["Golf","25.63","23.6","0","49.23"],
				["Tee","35.56","24.69","5555","5615.25"],
				["Grand Total","","","",""]
				];
			let runTotal_compact2 = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","36","73.13","73.13",""],
				["Fancy","11.96","25.7","25.7",""],
				["Golf","13","25.12","25.12",""],
				["Tee","11.04","22.31","22.31",""],
				["North","13.08","13.08","5568.08",""],
				["Tee","13.08","13.08","5568.08",""],
				["West","36.13","61.03","61.03",""],
				["Fancy","12.06","12.06","12.06",""],
				["Golf","12.63","24.11","24.11",""],
				["Tee","11.44","24.86","24.86",""],
				["Grand Total","85.21","147.24","5702.24",""]
				];
			let runTotal_stdDev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","0.980612054","1.254843948","0","1.028132611"],
				["Fancy","#DIV/0!","#DIV/0!","0","1.258650071"],
				["Golf","#DIV/0!","#DIV/0!","0","0.622253967"],
				["Tee","#DIV/0!","#DIV/0!","0","0.16263456"],
				["North","#DIV/0!","1.254843948","#DIV/0!","3919.757345"],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!","3918.891847"],
				["West","#DIV/0!","2.626631103","#DIV/0!","3920.592318"],
				["Fancy","#DIV/0!","#DIV/0!","0","#DIV/0!"],
				["Golf","#DIV/0!","#DIV/0!","0","1.435426766"],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!","3920.291919"],
				["Grand Total","","","",""]
				];
			let runTotal_stdDev_compact2 = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","0.980612054","2.235456002","2.235456002",""],
				["Fancy","#DIV/0!","#DIV/0!","#DIV/0!",""],
				["Golf","#DIV/0!","#DIV/0!","#DIV/0!",""],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!",""],
				["North","#DIV/0!","#DIV/0!","#DIV/0!",""],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!",""],
				["West","0.595175044","1.9669622","1.9669622",""],
				["Fancy","#DIV/0!","#DIV/0!","#DIV/0!",""],
				["Golf","#DIV/0!","#DIV/0!","#DIV/0!",""],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!",""],
				["Grand Total","0.774009351","1.896230364","#DIV/0!",""]
				];
			let differenceNext_stdDev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","#DIV/0!","1.254843948","#DIV/0!","-3917.70108"],
				["Fancy","#DIV/0!","#DIV/0!","0","#DIV/0!"],
				["Golf","#DIV/0!","#DIV/0!","0","-0.190918831"],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!","-3918.566578"],
				["North","#DIV/0!","-1.371787156","#DIV/0!","3917.89424"],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!","3917.329141"],
				["West","","","",""],
				["Fancy","","","",""],
				["Golf","","","",""],
				["Tee","","","",""],
				["Grand Total","","","",""]
				];
			let percentDiffNext_stdDev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","","","#NULL!","-0.999737636"],
				["Fancy","#DIV/0!","#DIV/0!","#NULL!",""],
				["Golf","#DIV/0!","#DIV/0!","#NULL!","-0.234782609"],
				["Tee","#DIV/0!","#DIV/0!","#NULL!","-0.999958498"],
				["North","#DIV/0!","#NULL!","#DIV/0!","4692.240335"],
				["Tee","#DIV/0!","#NULL!","#DIV/0!","2797.949495"],
				["West","","","",""],
				["Fancy","","","",""],
				["Golf","","","",""],
				["Tee","","","",""],
				["Grand Total","","","",""]
				]
			let percentNext_stdDev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","","","#NULL!","0.000262364"],
				["Fancy","#DIV/0!","#DIV/0!","#NULL!",""],
				["Golf","#DIV/0!","#DIV/0!","#NULL!","0.765217391"],
				["Tee","#DIV/0!","#DIV/0!","#NULL!","4.15019E-05"],
				["North","#DIV/0!","#NULL!","#DIV/0!","4693.240335"],
				["Tee","#DIV/0!","#NULL!","#DIV/0!","2798.949495"],
				["West","1","1","","1"],
				["Fancy","","","",""],
				["Golf","","","","1"],
				["Tee","","","","1"],
				["Grand Total","","","",""]
				];
			let percentRunTotal_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","0.422485624","0.598581332","0","0.012824785"],
				["Fancy","0.497918401","1","","0.680614407"],
				["Golf","0.507218104","0.513559322","","0.510257973"],
				["Tee","0.310461192","0.456460105","0","0.003973109"],
				["North","0.575988734","0.598581332","1","0.989297188"],
				["Tee","0.678290214","0.456460105","1","0.995572771"],
				["West","1","1","1","1"],
				["Fancy","1","1","","1"],
				["Golf","1","1","","1"],
				["Tee","1","1","1","1"],
				["Grand Total","","","",""]
				];
			let percentRunTotal_compact2 = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","0.492274033","1","1",""],
				["Fancy","0.46536965","1","1",""],
				["Golf","0.517515924","1","1",""],
				["Tee","0.494845361","1","1",""],
				["North","0.002349104","0.002349104","1",""],
				["Tee","0.002349104","0.002349104","1",""],
				["West","0.592003932","1","1",""],
				["Fancy","1","1","1",""],
				["Golf","0.523849025","1","1",""],
				["Tee","0.460176991","1","1",""],
				["Grand Total","0.01494325","0.025821432","1",""]
				];
			let percentRunTotal_stdDev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","1.266925331","1.118178981","#DIV/0!","0.000668801"],
				["Fancy","#DIV/0!","#DIV/0!","","#DIV/0!"],
				["Golf","#DIV/0!","#DIV/0!","","0.433497537"],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!","4.14853E-05"],
				["North","#DIV/0!","1.118178981","#DIV/0!","2.549805584"],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!","0.999642866"],
				["West","#DIV/0!","2.340564893","#DIV/0!","2.550348735"],
				["Fancy","#DIV/0!","#DIV/0!","","#DIV/0!"],
				["Golf","#DIV/0!","#DIV/0!","","1"],
				["Tee","#DIV/0!","#DIV/0!","#DIV/0!","1"],
				["Grand Total","","","",""]
				];
			let rankAscending_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","2","2","","2"],
				["Fancy","1","1","","2"],
				["Golf","2","2","","2"],
				["Tee","1","1","","1"],
				["North","1","","1","3"],
				["Tee","3","","1","3"],
				["West","3","1","","1"],
				["Fancy","2","","","1"],
				["Golf","1","1","","1"],
				["Tee","2","2","","2"],
				["Grand Total","","","",""]
				];
			let rankAscending_compact2 = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","1","2","",""],
				["Fancy","1","2","",""],
				["Golf","2","1","",""],
				["Tee","1","2","",""],
				["North","1","","2",""],
				["Tee","1","","2",""],
				["West","2","1","",""],
				["Fancy","1","","",""],
				["Golf","2","1","",""],
				["Tee","1","2","",""],
				["Grand Total","2","1","3",""]
				];
			let rankAscending_stdDev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","2","1","","2"],
				["Fancy","","","","1"],
				["Golf","","","","1"],
				["Tee","","","","1"],
				["North","","","","3"],
				["Tee","","","","3"],
				["West","1","2","","1"],
				["Fancy","","","",""],
				["Golf","","","","2"],
				["Tee","","","","2"],
				["Grand Total","","","",""]
				];
			let rankDescending_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","2","1","","2"],
				["Fancy","2","1","","1"],
				["Golf","1","1","","1"],
				["Tee","3","2","","3"],
				["North","3","","1","1"],
				["Tee","1","","1","1"],
				["West","1","2","","3"],
				["Fancy","1","","","2"],
				["Golf","2","2","","2"],
				["Tee","2","1","","2"],
				["Grand Total","","","",""]
				];
			let rankDescending_compact2 = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","2","1","",""],
				["Fancy","2","1","",""],
				["Golf","1","2","",""],
				["Tee","2","1","",""],
				["North","2","","1",""],
				["Tee","2","","1",""],
				["West","1","2","",""],
				["Fancy","1","","",""],
				["Golf","1","2","",""],
				["Tee","2","1","",""],
				["Grand Total","2","3","1",""]
				];
			let rankDescending_stdDev_compact = [
				["Sum of Price","Column Labels","","",""],
				["Row Labels","Boy","Girl","Helicopter","Grand Total"],
				["East","1","2","","2"],
				["Fancy","","","","1"],
				["Golf","","","","2"],
				["Tee","","","","3"],
				["North","","","","1"],
				["Tee","","","","1"],
				["West","2","1","","3"],
				["Fancy","","","",""],
				["Golf","","","","1"],
				["Tee","","","","2"],
				["Grand Total","","","",""]
				];
			var pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 2);
			pivot.asc_addColField(api, 1);
			pivot.asc_addDataField(api, 5);
			function testShowAs(pivot, showAs, baseField, baseItem, subtotalType, standard, message) {
				return checkHistoryOperation(assert, pivot, standard, message, function(){
					var dataField = pivot.asc_getDataFields()[0];
					let props = new Asc.CT_DataField();
					props.setShowAs(showAs, baseField, baseItem);
					props.asc_setSubtotal(subtotalType)
					dataField.asc_set(api, pivot, 0, props);
				});
			}
			AscCommon.History.Clear();
			setPivotLayout(pivot, 'compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfTotal, 0, 0, c_oAscDataConsolidateFunction.Sum, percentOfTotal_compact, 'percentOfTotal_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Difference, 0, AscCommonExcel.st_BASE_ITEM_NEXT, c_oAscDataConsolidateFunction.Sum, differenceNext_compact, 'differenceNext_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Difference, 1, AscCommonExcel.st_BASE_ITEM_NEXT, c_oAscDataConsolidateFunction.Sum, differenceNext_compact2, 'differenceNext_compact2');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Difference, 0, AscCommonExcel.st_BASE_ITEM_PREV, c_oAscDataConsolidateFunction.Sum, differencePrev_compact, 'differencePrev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Difference, 1, AscCommonExcel.st_BASE_ITEM_PREV, c_oAscDataConsolidateFunction.Sum, differencePrev_compact2, 'differencePrev_compact2');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Difference, 0, 1, c_oAscDataConsolidateFunction.Sum, differenceBase_compact, 'differenceBase_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Difference, 1, 1, c_oAscDataConsolidateFunction.Sum, differenceBase_compact2, 'differenceBase_compact2');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfCol, 0, 0, c_oAscDataConsolidateFunction.Sum, percentOfCol_compact, 'percentOfCol_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfRow, 0, 0, c_oAscDataConsolidateFunction.Sum, percentOfRow_compact, 'percentOfRow_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Index, 0, 0, c_oAscDataConsolidateFunction.Sum, index_compact, 'index_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfParentRow, 0, 0, c_oAscDataConsolidateFunction.Sum, percentOfParentRow_compact, 'percentOfParentRow_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfParentCol, 0, 0, c_oAscDataConsolidateFunction.Sum, percentOfParentCol_compact, 'percentOfParentCol_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfParent, 0, 0, c_oAscDataConsolidateFunction.Sum, percentOfParent_compact, 'percentOfParent_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Percent, 0, AscCommonExcel.st_BASE_ITEM_NEXT, c_oAscDataConsolidateFunction.Sum, percentNext_compact, 'percentNext_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Percent, 0, AscCommonExcel.st_BASE_ITEM_PREV, c_oAscDataConsolidateFunction.Sum, percentPrev_compact, 'percentPrev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Percent, 0, 1, c_oAscDataConsolidateFunction.Sum, percentBase_compact, 'percentBase_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentDiff, 0, AscCommonExcel.st_BASE_ITEM_NEXT, c_oAscDataConsolidateFunction.Sum, percentDiffNext_compact, 'percentDiffNext_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentDiff, 0, AscCommonExcel.st_BASE_ITEM_PREV, c_oAscDataConsolidateFunction.Sum, percentDiffPrev_compact, 'percentDiffPrev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentDiff, 0, 1, c_oAscDataConsolidateFunction.Sum, percentDiffBase_compact, 'percentDiffBase_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RunTotal, 0, 0, c_oAscDataConsolidateFunction.Sum, runTotal_compact, 'runTotal_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RunTotal, 1, 0, c_oAscDataConsolidateFunction.Sum, runTotal_compact2, 'runTotal_compact2');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RunTotal, 0, 0, c_oAscDataConsolidateFunction.StdDev, runTotal_stdDev_compact, 'runTotal_stdDev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RunTotal, 1, 0, c_oAscDataConsolidateFunction.StdDev, runTotal_stdDev_compact2, 'runTotal_stdDev_compact2');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Difference, 0, AscCommonExcel.st_BASE_ITEM_NEXT, c_oAscDataConsolidateFunction.StdDev, differenceNext_stdDev_compact, 'differenceNext_stdDev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentDiff, 0, AscCommonExcel.st_BASE_ITEM_NEXT, c_oAscDataConsolidateFunction.StdDev, percentDiffNext_stdDev_compact, 'percentDiffNext_stdDev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Percent, 0, AscCommonExcel.st_BASE_ITEM_NEXT, c_oAscDataConsolidateFunction.StdDev, percentNext_stdDev_compact, 'percentNext_stdDev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfRunningTotal, 0, 0, c_oAscDataConsolidateFunction.Sum, percentRunTotal_compact, 'percentRunTotal_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfRunningTotal, 1, 0, c_oAscDataConsolidateFunction.Sum, percentRunTotal_compact2, 'percentRunTotal_compact2');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.PercentOfRunningTotal, 0, 0, c_oAscDataConsolidateFunction.StdDev, percentRunTotal_stdDev_compact, 'percentRunTotal_stdDev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RankAscending, 0, 0, c_oAscDataConsolidateFunction.Sum, rankAscending_compact, 'rankAscending_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RankAscending, 1, 0, c_oAscDataConsolidateFunction.Sum, rankAscending_compact2, 'rankAscending_compact2');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RankAscending, 0, 0, c_oAscDataConsolidateFunction.StdDev, rankAscending_stdDev_compact, 'rankAscending_stdDev_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RankDescending, 0, 0, c_oAscDataConsolidateFunction.Sum, rankDescending_compact, 'rankDescending_compact');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RankDescending, 1, 0, c_oAscDataConsolidateFunction.Sum, rankDescending_compact2, 'rankDescending_compact2');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.RankDescending, 0, 0, c_oAscDataConsolidateFunction.StdDev, rankDescending_stdDev_compact, 'rankDescending_stdDev_compact');
			setPivotLayout(pivot, 'tabular');
			pivot = testShowAs(pivot, Asc.c_oAscShowDataAs.Difference, 0, AscCommonExcel.st_BASE_ITEM_NEXT, c_oAscDataConsolidateFunction.Sum, differenceNext_tabular, 'differenceNext_tabular', 0);
			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}
	function testPivotShowDetails() {
		QUnit.test('Test: Show Details', function (assert) {
			const testData =  [
				["Region","Gender","Style","Ship date","Units","Price","Cost"],
				["East","Boy","Tee","1","12","11.04","10.42"],
				["East","Boy","Golf","1","12","13","12.6"],
				["East","Boy","Fancy","2","12","11.96","11.74"],
				["East","Girl","Tee","2","10","11.27","10.56"],
				["East","Girl","Golf","1","10","12.12","11.95"],
				["East","Girl","Fancy","2","10","13.74","13.33"],
				["West","Boy","Tee","1","11","11.44","10.94"],
				["West","Boy","Golf","2","11","12.63","11.73"],
				["West","Boy","Fancy","1","11","12.06","11.51"],
				["West","Girl","Tee","2","15","13.42","13.29"],
				["West","Girl","Golf","1","15","11.48","10.67"]
			];
			const standardNoFilterEastGT = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['East', 'Boy', 'Tee', '1', '12', '11.04', '10.42'],
				['East', 'Boy', 'Golf', '1', '12', '13', '12.6'],
				['East', 'Boy', 'Fancy', '2', '12', '11.96', '11.74'],
				['East', 'Girl', 'Tee', '2', '10', '11.27', '10.56'],
				['East', 'Girl', 'Golf', '1', '10', '12.12', '11.95'],
				['East', 'Girl', 'Fancy', '2', '10', '13.74', '13.33'],
			];
			const standardNoFilterFancyGT = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['East', 'Boy', 'Fancy', '2', '12', '11.96', '11.74'],
				['East', 'Girl', 'Fancy', '2', '10', '13.74', '13.33'],
			];
			const standardNoFilter10GT = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['East', 'Girl', 'Fancy', '2', '10', '13.74', '13.33'],
			];
			const standardNoFilterEastGirl = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['East', 'Girl', 'Tee', '2', '10', '11.27', '10.56'],
				['East', 'Girl', 'Golf', '1', '10', '12.12', '11.95'],
				['East', 'Girl', 'Fancy', '2', '10', '13.74', '13.33'],
			];
			const standardNoFilterFancyGirl = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['East', 'Girl', 'Fancy', '2', '10', '13.74', '13.33'],
			];
			const standardNoFilter12Girl = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['', '', '', '', '', '', ''],
			];
			const standardFilterEastGT = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['East', 'Boy', 'Tee', '1', '12', '11.04', '10.42'],
				['East', 'Boy', 'Golf', '1', '12', '13', '12.6'],
				['East', 'Girl', 'Golf', '1', '10', '12.12', '11.95'],
			];
			const standardFilterGTGT = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['East', 'Boy', 'Tee', '1', '12', '11.04', '10.42'],
				['East', 'Boy', 'Golf', '1', '12', '13', '12.6'],
				['East', 'Girl', 'Golf', '1', '10', '12.12', '11.95'],
				['West', 'Boy', 'Tee', '1', '11', '11.44', '10.94'],
				['West', 'Boy', 'Fancy', '1', '11', '12.06', '11.51'],
				['West', 'Girl', 'Golf', '1', '15', '11.48', '10.67'],
			];
			const standardGroupFilter = [
				['Region', 'Gender', 'Style', 'Ship date', 'Units', 'Price', 'Cost'],
				['East', 'Boy', 'Golf', '1', '12', '13', '12.6'],
				['East', 'Girl', 'Golf', '1', '10', '12.12', '11.95'],
			];
			function testPivotCellForDetails(assert, pivot, row, col, standard, message) {
				let undoStandard = [];
				for (let i = 0; i < standard.length; i += 1) {
					undoStandard[i] = [];
					undoStandard[i].length = standard[0].length;
					undoStandard[i].fill("");
				}
				let res = checkHistoryOperation2(assert, pivot, standard, message, undoStandard, function () {
					const indexes = pivot.getItemsIndexesByActiveCell(row, col);
					const arrayItemFieldsMap = pivot.getNoFilterItemFieldsMapArray(indexes.rowItemIndex, indexes.colItemIndex)
					pivot.showDetails(wsDetails, arrayItemFieldsMap);
				}, function (assert, pivot, standard, message) {
					let cells = [];
					for (let i = 0; i < standard.length; i += 1) {
						cells[i] = [];
						for (let j = 0; j < standard[0].length; j += 1) {
							cells[i][j] = wsDetails.getCell3(i, j).getValue();
						}
					}
					assert.deepEqual(cells, standard, message)
				});
				wsDetails.removeRows(0, wsDetails.getRowsCount());
				return res;
			}
			function getNewFilter(fld, index) {
				var autoFilterObject = new Asc.AutoFiltersOptions();
				pivot.fillAutoFiltersOptions(autoFilterObject, fld);
				for (var i = 0; i < autoFilterObject.values.length; ++i) {
					autoFilterObject.values[i].visible = i == index;
				}
				autoFilterObject.filter.type = Asc.c_oAscAutoFilterTypes.Filters;
				return autoFilterObject;
			};
			let testDataRange = new Asc.Range(0, 0, testData[0].length - 1, testData.length - 1);
			fillData(wsData, testData, testDataRange);
			let dataRef = wsData.getName() + "!" + testDataRange.getName();
			let pivot = api._asc_insertPivot(wb, dataRef, ws, reportRange);
			pivot.asc_getStyleInfo().asc_setName(api, pivot, pivotStyle);
			pivot.asc_addRowField(api, 0);
			pivot.asc_addRowField(api, 2);
			pivot.asc_addRowField(api, 4);
			pivot.asc_addColField(api, 1);
			pivot.asc_addDataField(api, 5);

			AscCommon.History.Clear();
			pivot = testPivotCellForDetails(assert, pivot, 4, 3, standardNoFilterEastGT, 'no-filter East | GT');
			pivot = testPivotCellForDetails(assert, pivot, 5, 3, standardNoFilterFancyGT, 'no-filter East -> Fancy | GT');
			pivot = testPivotCellForDetails(assert, pivot, 6, 3, standardNoFilter10GT, 'no-filter East -> Fancy -> 10 (Units) | GT');
			pivot = testPivotCellForDetails(assert, pivot, 4, 2, standardNoFilterEastGirl, 'no-filter East | Girl');
			pivot = testPivotCellForDetails(assert, pivot, 5, 2, standardNoFilterFancyGirl, 'no-filter East -> Fancy | Girl');
			pivot = testPivotCellForDetails(assert, pivot, 5, 2, standardNoFilterFancyGirl, 'no-filter East -> Fancy | Girl');
			pivot = testPivotCellForDetails(assert, pivot, 7, 2, standardNoFilter12Girl, 'no-filter East -> Fancy -> 12 (Units) | Girl');

			pivot.asc_addPageField(api, 3);
			pivot.filterByFieldIndex(api, getNewFilter(3, 0), 3, true);

			AscCommon.History.Clear();
			pivot = testPivotCellForDetails(assert, pivot, 4, 3, standardFilterEastGT, 'filter 1 (ship date) East | GT');
			pivot = testPivotCellForDetails(assert, pivot, 17, 3, standardFilterGTGT, 'filter 1 (ship date) GTGT');

			const group = new PivotLayoutGroup();
			group.fld = 4;
			group.groupMap = {
				0: 1,
				2: 1
			};
			const onRepeat = function () {
				api._groupPivot(true, onRepeat);
			}
			pivot.groupPivot(api, group, false, onRepeat);
			pivot = testPivotCellForDetails(assert, pivot, 6, 3, standardGroupFilter, 'filter 1 (ship date) Group 1 Units (10, 12)');
			ws.deletePivotTables(new AscCommonExcel.MultiplyRange(pivot.getReportRanges()).getUnionRange());
		});
	}
	QUnit.module("Pivot");

	function startTests() {
		QUnit.start();

		testValidations();

		testLayout("compact");

		testLayout("outline");

		testLayout("tabular");

		testLayout("gridDropZones");

		testLayout("showHeaders");

		testLayoutValues("compact");

		testLayoutValues("outline");

		testLayoutValues("tabular");

		testLayoutSubtotal("compact");

		testLayoutSubtotal("tabular");

		testPivotInsertBlankRow();

		testPivotPageFilterLayout();

		testFieldProperty();

		testFieldSubtotal();

		testDataValues();

		testHeaderRename();

		testPivotManipulationField();

		testPivotManipulationValues();

		testDataRefresh();

		testDataSource();

		testFiltersValueFilter();

		testFiltersValueFilterBug46141();

		//todo FilterVal is set after History.Add
		//testFiltersTop10();

		testFiltersLabel();

		testFiltersReIndex();

		testNumFormat();

		testPivotMisc();

		testPivotShowAs();

		testPivotShowDetails();
	}
});
