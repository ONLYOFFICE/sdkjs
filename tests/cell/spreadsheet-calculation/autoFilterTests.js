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

$(function () {

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
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable = function () {
	};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function(){};
	cDate.prototype.getCurrentDate = function () {
		return new cDate(2023, 4, 15, 0, 0, 0);
	}


	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function() {}
	};
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api._openDocument(AscCommon.getEmpty());
	api._openOnClient();
	api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
	var wb = api.wbModel;
	wb.handlers.add("getSelectionState", function () {
		return null;
	});
	wb.handlers.add("getLockDefNameManagerStatus", function () {
		return true;
	});
	api.wb.cellCommentator = new AscCommonExcel.CCellCommentator({
		model: api.wbModel.aWorksheets[0],
		collaborativeEditing: null,
		draw: function() {
		},
		handlers: {
			trigger: function() {
				return false;
			}
		}
	});

	AscCommonExcel.CCellCommentator.prototype.isLockedComment = function (oComment, callbackFunc) {
		callbackFunc(true);
	};
	AscCommonExcel.CCellCommentator.prototype.drawCommentCells = function () {
	};
	AscCommonExcel.CCellCommentator.prototype.ascCvtRatio = function () {
	};

	var wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();
	var ws = api.wbModel.aWorksheets[0];

	var getRange = function (c1, r1, c2, r2) {
		return new window["Asc"].Range(c1, r1, c2, r2);
	};
	const getRangeWithData = function (ws, data) {
		let range = ws.getRange4(0, 0);

		range.fillData(data);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];

		return ws;
	}
	const createDynamicFilter = function (ws, filterType, colId) {
		// Initialization filter option and dynamic filter
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: colId, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Imitate choose filter option
		dynamicFilter.asc_setType(filterType)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		return ws.autoFilters;
	}
	const clearData = function (c1, r1, c2, r2) {
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	}

	QUnit.test("Test: \"simple tests\"", function (assert) {
		let testData = [
			["test1", "test2"],
			["", "44851"],
			["closed", ""],
			["closed", ""],
			["d", "44852"],
			["closed", ""],
			["", "44851"],
			["", "44851"],
			["closed", "44851"]
		];

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, "check filter range");
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, "check filter range");
		assert.strictEqual(ws.AutoFilter.Ref.r2, 8, "check filter range");
		assert.strictEqual(ws.AutoFilter.Ref.c2, 1, "check filter range");

		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		autoFiltersOptions.values[0].asc_setVisible(false);//hide "closed"
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.Filters);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//2,3,5,8 hidden
		assert.strictEqual(ws.getRowHidden(1), false, "check filter hidden values");
		assert.strictEqual(ws.getRowHidden(2), true, "check filter hidden values");
		assert.strictEqual(ws.getRowHidden(3), true, "check filter hidden values");
		assert.strictEqual(ws.getRowHidden(4), false, "check filter hidden values");
		assert.strictEqual(ws.getRowHidden(5), true, "check filter hidden values");
		assert.strictEqual(ws.getRowHidden(6), false, "check filter hidden values");
		assert.strictEqual(ws.getRowHidden(7), false, "check filter hidden values");
		assert.strictEqual(ws.getRowHidden(8), true, "check filter hidden values");

		autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 1, id: null});
		autoFiltersOptions.values[0].asc_setVisible(false);//hide "44851"
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.Filters);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		assert.strictEqual(ws.getRowHidden(1), true, "check filter hidden values_2");
		assert.strictEqual(ws.getRowHidden(2), true, "check filter hidden values_2");
		assert.strictEqual(ws.getRowHidden(3), true, "check filter hidden values_2");
		assert.strictEqual(ws.getRowHidden(4), false, "check filter hidden values_2");
		assert.strictEqual(ws.getRowHidden(5), true, "check filter hidden values_2");
		assert.strictEqual(ws.getRowHidden(6), true, "check filter hidden values_2");
		assert.strictEqual(ws.getRowHidden(7), true, "check filter hidden values_2");
		assert.strictEqual(ws.getRowHidden(8), true, "check filter hidden values_2");

		ws.setRowHidden(false, 0, 8);

		assert.strictEqual(ws.getRowHidden(1), false, "check hidden row");
		assert.strictEqual(ws.getRowHidden(2), false, "check hidden row");
		assert.strictEqual(ws.getRowHidden(3), false, "check hidden row");
		assert.strictEqual(ws.getRowHidden(4), false, "check hidden row");
		assert.strictEqual(ws.getRowHidden(5), false, "check hidden row");
		assert.strictEqual(ws.getRowHidden(6), false, "check hidden row");
		assert.strictEqual(ws.getRowHidden(7), false, "check hidden row");
		assert.strictEqual(ws.getRowHidden(8), false, "check hidden row");

		autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 1, id: null});
		autoFiltersOptions.values[0].asc_setVisible(false);//hide "44851"
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.Filters);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		assert.strictEqual(ws.getRowHidden(1), true, "check filter hidden values_3");
		assert.strictEqual(ws.getRowHidden(2), true, "check filter hidden values_3");
		assert.strictEqual(ws.getRowHidden(3), true, "check filter hidden values_3");
		assert.strictEqual(ws.getRowHidden(4), false, "check filter hidden values_3");
		assert.strictEqual(ws.getRowHidden(5), true, "check filter hidden values_3");
		assert.strictEqual(ws.getRowHidden(6), true, "check filter hidden values_3");
		assert.strictEqual(ws.getRowHidden(7), true, "check filter hidden values_3");
		assert.strictEqual(ws.getRowHidden(8), true, "check filter hidden values_3");

		//Clearing data of sheet
		clearData(0,0,1,8);
	});
	QUnit.test('Test: "Date Filter - Today"', function (assert) {
		const testData = [
			['Dates'],
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'] // 16.05.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 3, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Today"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.today,0);

		// Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 15.05.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 16.05.2023 tomorrow must be hidden');

		// Clearing data of sheet
		clearData(0, 0, 0, 3);
	});
	QUnit.test('Test: "Date Filter - Yesterday"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'] // 15.05.2023 today
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 3, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Yesterday"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.yesterday, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 14.05.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 3);
	});
	QUnit.test('Test: "Date Filter - Tomorrow"', function (assert) {
		const testData = [
			['Dates'],
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'] // 17.05.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 3, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Tomorrow"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.tomorrow, 0);

		//Checking work of filter

		assert.strictEqual(ws.getRowHidden(1), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 16.05.2023 tomorrow must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 17.05.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 3);
	});
	QUnit.test('Test: "Date Filter - This week"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023 - 6 day of week
			['45060'], // 14.05.2023 - 0 day of week
			['45061'], // 15.05.2023 - 1 day of week
			['45062'], // 16.05.2023 - 2 day of week
			['45063'], // 17.05.2023 - 3 day of week
			['45064'], // 18.05.2023 - 4 day of week
			['45065'], // 19.05.2023 - 5 day of week
			['45066'], // 20.05.2023 - 6 day of week
			['45067']  // 21.05.2023 - 0 day of week
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 9, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "This week"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.thisWeek, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 14.05.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 15.05.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 16.05.2023 tomorrow must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), false, 'Value 17.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(6), false, 'Value 18.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(7), false, 'Value 19.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(8), false, 'Value 20.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 21.05.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 9);
	});
	QUnit.test('Test: "Date Filter - Next week"', function (assert) {
		const testData = [
			['Dates'],
			['45066'], // 20.05.2023 - 6 day of week
			['45067'], // 21.05.2023 - 0 day of week
			['45068'], // 22.05.2023 - 1 day of week
			['45069'], // 23.05.2023 - 2 day of week
			['45070'], // 24.05.2023 - 3 day of week
			['45071'], // 25.05.2023 - 4 day of week
			['45072'], // 26.05.2023 - 5 day of week
			['45073'], // 27.05.2023 - 6 day of week
			['45074']  // 28.05.2023 - 0 day of week
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData)
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 9, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Next week"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.nextWeek, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 20.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 21.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 22.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 23.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), false, 'Value 24.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(6), false, 'Value 25.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(7), false, 'Value 26.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(8), false, 'Value 27.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 28.05.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 9);
	});
	QUnit.test('Test: "Date Filter - Last week"', function (assert) {
		const testData = [
			['Dates'],
			['45052'], // 06.05.2023 - 6 day of week
			['45053'], // 07.05.2023 - 0 day of week
			['45054'], // 08.05.2023 - 1 day of week
			['45055'], // 09.05.2023 - 2 day of week
			['45056'], // 10.05.2023 - 3 day of week
			['45057'], // 11.05.2023 - 4 day of week
			['45058'], // 12.05.2023 - 5 day of week
			['45059'], // 13.05.2023 - 6 day of week
			['45060']  // 14.05.2023 - 0 day of week
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 9, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Last week"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.lastWeek, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 06.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 07.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 08.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 09.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), false, 'Value 10.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(6), false, 'Value 11.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(7), false, 'Value 12.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(8), false, 'Value 13.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 14.04.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 9);
	});
	QUnit.test('Test: "Date Filter - Last month"', function (assert) {
		const testData = [
			['Dates'],
			['45016'], // 31.03.2023
			['45017'], // 01.04.2023
			['45046'], // 30.04.2023
			['45047']  // 01.05.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Last month"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.lastMonth, 0);


		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.03.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.05.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - This month"', function (assert) {
		const testData = [
			['Dates'],
			['45046'], // 30.04.2023
			['45047'], // 01.05.2023
			['45077'], // 31.05.2023
			['45078'] // 01.06.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "This month"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.thisMonth, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.05.2023 yesterday not must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.06.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - Next month"', function (assert) {
		const testData = [
			['Dates'],
			['45077'], // 31.05.2023
			['45078'], // 01.06.2023
			['45107'], // 30.06.2023
			['45108']  // 01.07.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Next month"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.nextMonth, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.06.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.06.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.07.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - Next quarter"', function (assert) {
		const testData = [
			['Dates'],
			['45107'], // 30.06.2023
			['45108'], // 01.07.2023
			['45199'], // 30.09.2023
			['45200']  // 01.10.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Next quarter"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.nextQuarter, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.07.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.09.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.10.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - This quarter"', function (assert) {
		const testData = [
			['Dates'],
			['45016'], // 31.03.2023
			['45017'], // 01.04.2023
			['45107'], // 30.06.2023
			['45108'] // 01.07.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "This quarter"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.thisQuarter, 0);

		//Checking work of filter

		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.03.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.06.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.07.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - Last quarter"', function (assert) {
		const testData = [
			['Dates'],
			['44926'], // 31.12.2022
			['44927'], // 01.01.2023
			['45016'], // 31.03.2023
			['45017']  // 01.04.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Last quarter"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.lastQuarter, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2022 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.03.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.04.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - This year"', function (assert) {
		const testData = [
			['Dates'],
			['44926'], // 31.12.2022
			['44927'], // 01.01.2023
			['45291'], // 31.12.2023
			['45292']  // 01.01.2024

		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "В этом году"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.thisYear, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2022 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.2024 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - Next year"', function (assert) {
		const testData = [
			['Dates'],
			['45291'], // 31.12.2023
			['45292'], // 01.01.2024
			['45657'], // 31.12.2025
			['45658'] // 01.01.2025
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Next year"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.nextYear, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2024 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.2024 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.2025 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - Last year"', function (assert) {
		const testData = [
			['Dates'],
			['44561'], // 31.12.2021
			['44562'], // 01.01.2022
			['44926'], // 31.12.2022
			['44927'] // 01.01.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData)
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "В прошлом году"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.lastYear, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2021 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2022 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.2022 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - Year to date"', function (assert) {
		const testData = [
			['Dates'],
			['44926'], // 31.12.2022
			['44927'], // 01.01.2023
			['45068'], // 22.05.2023
			['45069'] // 23.05.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Year to Date"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.yearToDate, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2022 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 22.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 23.05.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> January"', function (assert) {
		const testData = [
			['Dates'],
			['364'], // 30.12.1900
			['365'], // 31.12.1900
			['0'],  // 01.01.1900
			['30'], // 31.01.1900
			['31'], // 01.02.1900
			['32'] // 02.02.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 6, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "January"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m1, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.12.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 31.12.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 01.01.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 31.01.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 01.02.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 02.02.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 6)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> February"', function (assert) {
		const testData = [
			['Dates'],
			['30'], // 31.01.1900
			['31'], // 01.02.1900
			['58'], // 28.02.1900
			['59'] // 01.03.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "February"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m2, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.01.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.02.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 28.02.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.03.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> March"', function (assert) {
		const testData = [
			['Dates'],
			['58'], // 28.02.1900
			['59'], // 01.03.1900
			['90'], // 31.03.1900
			['91'] // 01.04.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "March"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m3, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 28.02.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.03.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.03.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.04.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> April"', function (assert) {
		const testData = [
			['Dates'],
			['90'], // 31.03.1900
			['91'], // 01.04.1900
			['120'], // 30.04.1900
			['121'] // 01.05.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "April"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m4, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.03.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.04.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.04.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.05.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> May"', function (assert) {
		const testData = [
			['Dates'],
			['120'], // 30.04.1900
			['121'], // 01.05.1900
			['151'], // 31.05.1900
			['152'] // 01.06.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "May"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m5, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.04.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.05.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.05.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.06.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> June"', function (assert) {
		const testData = [
			['Dates'],
			['151'], // 31.05.1900
			['152'], // 01.06.1900
			['181'], // 30.06.1900
			['182'] // 01.07.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "June"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m6, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.05.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.06.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.06.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.07.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> July"', function (assert) {
		const testData = [
			['Dates'],
			['181'], // 30.06.1900
			['182'], // 01.07.1900
			['212'], // 31.07.1900
			['213'] // 01.08.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "July"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m7, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.06.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.07.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.07.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.08.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> August"', function (assert) {
		const testData = [
			['Dates'],
			['212'], // 31.07.1900
			['213'], // 01.08.1900
			['243'], // 31.08.1900
			['244']  // 01.09.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "August"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m8, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.07.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.08.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.08.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.09.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> September"', function (assert) {
		const testData = [
			['Dates'],
			['243'], // 31.08.1900
			['244'], // 01.09.1900
			['273'], // 30.09.1900
			['274']  // 01.10.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "September"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m9, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.08.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.09.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.09.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.10.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> October"', function (assert) {
		const testData = [
			['Dates'],
			['273'], // 30.09.1900
			['274'], // 01.10.1900
			['304'], // 31.10.1900
			['305']  // 01.11.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "October"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m10, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.09.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.10.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.10.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.11.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> November"', function (assert) {
		const testData = [
			['Dates'],
			['304'], // 31.10.1900
			['305'], // 01.11.1900
			['334'], // 30.11.1900
			['335']  // 01.12.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "November"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m11, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.10.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.11.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.11.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.12.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> December"', function (assert) {
		const testData = [
			['Dates'],
			['334'], // 30.11.1900
			['335'], // 01.12.1900
			['365'], // 31.12.1900
			['0']    // 01.01.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "December"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m12, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.11.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.12.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> Quarter 1"', function (assert) {
		const testData = [
			['Dates'],
			['365'], // 31.12.1900
			['0'], // 01.01.1900
			['90'], // 31.03.1900
			['91'] // 01.04.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Quarter 1"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.q1, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.03.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.04.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> Quarter 2"', function (assert) {
		const testData = [
			['Dates'],
			['90'], // 31.03.1900
			['91'], // 01.04.1900
			['181'], // 30.06.1900
			['182']  // 01.07.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Quarter 2"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.q2, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.03.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.04.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.06.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.07.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> Quarter 3"', function (assert) {
		const testData = [
			['Dates'],
			['181'], // 30.06.1900
			['182'], // 01.07.1900
			['273'], // 30.09.1900
			['274']  // 01.10.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Quarter 3"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.q3, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.06.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.07.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.09.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.10.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> Quarter 4"', function (assert) {
		const testData = [
			['Dates'],
			['273'], // 30.09.1900
			['274'], // 01.10.1900
			['365'], // 31.12.1900
			['0']    // 01.01.1900
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 4, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Imitate choosing filter "Quarter 4"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.q4, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.09.1900 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.10.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.1900 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.1900 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});

	QUnit.module("CopyPaste");
});