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
	Asc.ReadDefTableStyles = function () {
	};
	cDate.prototype.getCurrentDate = function () {
		return new cDate(2023, 4, 15, 0, 0, 0);
	};


	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
		}
	};
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api._openDocument(AscCommon.getEmpty());
	api._openOnClient();
	api.initCollaborativeEditing({});
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
		draw: function () {
		},
		handlers: {
			trigger: function () {
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
	};
	const createDynamicFilter = function (ws, filterType, colId) {
		// Initialization filter option and dynamic filter
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: colId, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Imitate choose filter option
		dynamicFilter.asc_setType(filterType);
		dynamicFilter.init(getRange(0, 0, 0, 0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		return ws.autoFilters;
	};

	const createCustomFilter = function (colId, aProps, and) {
		//apply filter
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: colId, id: null});
		let customFilters = new Asc.CustomFilters();
		customFilters.And = !!and;
		if (aProps && aProps.length) {
			for (let i = 0; i < aProps.length; i++) {
				let customFilter = new Asc.CustomFilter();
				customFilter.Operator = aProps[i].operator;
				customFilter.Val = aProps[i].val;
				if (!customFilters.CustomFilters) {
					customFilters.CustomFilters = [];
				}
				customFilters.CustomFilters.push(customFilter);
			}
		}

		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.CustomFilters);
		autoFiltersOptions.filter.asc_setFilter(customFilters);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);
	};

	const clearData = function (c1, r1, c2, r2) {
		ws.autoFilters.deleteAutoFilter(getRange(0, 0, 0, 0));
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	};

	const setFilterOptionsVisible = function (autoFiltersOptions, aVal, bDate) {
		if (aVal && autoFiltersOptions) {
			for (let i = 0; i < aVal.length; i++) {
				for (let j = 0; j < autoFiltersOptions.values.length; j++) {
					let elem = autoFiltersOptions.values[j];
					if (bDate && elem.year === aVal[i].year && elem.month === aVal[i].month && elem.day === aVal[i].day
						&& elem.hour === aVal[i].hour && elem.minute === aVal[i].minute) {
						elem.asc_setVisible(aVal[i].visible);
						elem.asc_setDateTimeGrouping(aVal[i].dateTimeGrouping);
						break;
					} else if (elem.text === aVal[i].text) {
						elem.asc_setVisible(aVal[i].visible);
						break;
					}
				}
			}
		}
	};

	const checkFilterRef = function (assert, r1, c1, r2, c2) {
		assert.strictEqual(ws.AutoFilter.Ref.r1, r1, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, c1, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, r2, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, c2, 'Check finish point column filter range');
	};

	const checkHiddenRows = function (assert, data, oHiddenRows, descPrefix) {
		if (!descPrefix) {
			descPrefix = "";
		}
		for (let i = 0; i < data.length; i++) {
			assert.strictEqual(ws.getRowHidden(i), !!oHiddenRows[i], descPrefix + 'Value ' + data[i] + ' must ' + (oHiddenRows[i] ? " " : " not ") + 'be hidden');
		}
	};

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

		checkFilterRef(assert, 0, 0, 8, 1);

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
		clearData(0, 0, 1, 8);
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
		checkFilterRef(assert, 0, 0, 3, 0);

		// Imitate choosing filter "Today"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.today, 0);

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
		checkFilterRef(assert, 0, 0, 3, 0);

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
		checkFilterRef(assert, 0, 0, 3, 0);

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
		checkFilterRef(assert, 0, 0, 9, 0);

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
		//TODO
		//assert.strictEqual(ws.getRowHidden(9), true, 'Value 21.05.2023 must be hidden');

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
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 9, 0);

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
		//TODO
		//assert.strictEqual(ws.getRowHidden(9), true, 'Value 28.05.2023 must be hidden');

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
		checkFilterRef(assert, 0, 0, 9, 0);

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
		//TODO
		//assert.strictEqual(ws.getRowHidden(9), true, 'Value 14.04.2023 must be hidden');

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
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Last month"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.lastMonth, 0);


		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.03.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.04.2023 must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.05.2023 must be hidden');

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
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "This month"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.thisMonth, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.05.2023 yesterday not must be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.06.2023 must be hidden');

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
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Next month"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.nextMonth, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.06.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.06.2023 must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.07.2023 must be hidden');

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
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Next quarter"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.nextQuarter, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.07.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.09.2023 must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.10.2023 must be hidden');

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
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "This quarter"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.thisQuarter, 0);

		//Checking work of filter

		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.03.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.06.2023 must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.07.2023 must be hidden');

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
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Last quarter"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.lastQuarter, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2022 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.03.2023 must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.04.2023 must be hidden');

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
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "В этом году"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.thisYear, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2022 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.2023 today must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.2024 must be hidden');

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
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Next year"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.nextYear, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2024 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.2024 must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.2025 must be hidden');

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
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "В прошлом году"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.lastYear, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2021 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2022 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.2022 must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Test: "Date Filter - Year to date"', function (assert) {
		const testData = [
			['Dates'],
			['44926'], // 31.12.2022
			['44927'], // 01.01.2023
			['45061'], // 15.05.2023
			['45062'] // 16.05.2023
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Year to Date"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.yearToDate, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2022 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 15.05.2023 must not be hidden');
		//TODO
		//assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	/*
		Cases with years 1900-1920 for filters "All Dates in the Period" don't work correct because dates move on 1-2 days
		 if you use method getDateFromExcel() for excelDate values.
		e.g.
		01.01.1900 (excelDate: 0) -> 30.12.1899
		01.01.1901 (excelDate: 366) -> 31.12.1900
		01.01.1910 (excelDate: 3653) -> 31.12.1909
		...
		01.01.1920 (excelDate: 7306) -> 01.01.1920
	*/
	QUnit.test('Test: "Date Filter - All Dates in the Period -> January"', function (assert) {
		const testData = [
			['Dates'],
			['36890'], // 30.12.2000
			['36891'], // 31.12.2000
			['36526'], // 01.01.2000
			['36556'], // 31.01.2000
			['36557'], // 01.02.2000
			['36558'] // 02.02.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 6, 0);

		// Imitate choosing filter "January"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m1, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.12.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 31.12.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 01.01.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 31.01.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 01.02.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 02.02.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 6)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> February"', function (assert) {

		const testData = [
			['Dates'],
			['36556'], // 31.01.2000
			['36557'], // 01.02.2000
			['36584'], // 28.02.2000
			['36586']  // 01.03.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "February"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m2, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.01.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.02.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 28.02.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.03.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> March"', function (assert) {
		const testData = [
			['Dates'],
			['36584'], // 28.02.2000
			['36586'], // 01.03.2000
			['36616'], // 31.03.2000
			['36617']  // 01.04.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "March"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m3, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 28.02.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.03.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.03.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.04.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> April"', function (assert) {
		const testData = [
			['Dates'],
			['36616'], // 31.03.2000
			['36617'], // 01.04.2000
			['36646'], // 30.04.2000
			['36647']  // 01.05.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "April"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m4, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.03.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.04.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.04.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.05.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> May"', function (assert) {
		const testData = [
			['Dates'],
			['36646'], // 30.04.2000
			['36647'], // 01.05.2000
			['36677'], // 31.05.2000
			['36678']  // 01.06.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "May"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m5, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.04.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.05.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.05.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.06.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> June"', function (assert) {
		const testData = [
			['Dates'],
			['36677'], // 31.05.2000
			['36678'], // 01.06.2000
			['36707'], // 30.06.2000
			['36708']  // 01.07.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "June"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m6, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.05.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.06.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.06.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.07.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> July"', function (assert) {
		const testData = [
			['Dates'],
			['36707'], // 30.06.2000
			['36708'], // 01.07.2000
			['36738'], // 31.07.2000
			['36739']  // 01.08.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "July"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m7, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.06.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.07.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.07.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.08.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> August"', function (assert) {
		const testData = [
			['Dates'],
			['36738'], // 31.07.2000
			['36739'], // 01.08.2000
			['36769'], // 31.08.2000
			['36770'] // 01.09.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "August"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m8, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.07.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.08.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.08.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.09.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> September"', function (assert) {
		const testData = [
			['Dates'],
			['36769'], // 31.08.2000
			['36770'], // 01.09.2000
			['36799'], // 30.09.2000
			['36800'], // 01.10.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "September"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m9, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.08.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.09.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.09.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.10.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> October"', function (assert) {
		const testData = [
			['Dates'],
			['36799'], // 30.09.2000
			['36800'], // 01.10.2000
			['36830'], // 31.10.2000
			['36831']  // 01.11.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "October"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m10, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.09.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.10.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.10.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.11.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> November"', function (assert) {
		const testData = [
			['Dates'],
			['36830'], // 31.10.2000
			['36831'], // 01.11.2000
			['36860'], // 30.11.2000
			['36861']  // 01.12.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "November"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m11, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.10.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.11.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.11.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.12.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> December"', function (assert) {
		const testData = [
			['Dates'],
			['36860'], // 30.11.2000
			['36861'], // 01.12.2000
			['36891'], // 31.12.2000
			['36526']  // 01.01.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "December"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.m12, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.11.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.12.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> Quarter 1"', function (assert) {
		const testData = [
			['Dates'],
			['36891'], // 31.12.2000
			['36526'], // 01.01.2000
			['36616'], // 31.03.2000
			['36617']  // 01.04.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Quarter 1"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.q1, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.12.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.01.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.03.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.04.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> Quarter 2"', function (assert) {
		const testData = [
			['Dates'],
			['36616'], // 31.03.2000
			['36617'], // 01.04.2000
			['36707'], // 30.06.2000
			['36708']  // 01.07.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Quarter 2"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.q2, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 31.03.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.04.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.06.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.07.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> Quarter 3"', function (assert) {
		const testData = [
			['Dates'],
			['36707'], // 30.06.2000
			['36708'], // 01.07.2000
			['36799'], // 30.09.2000
			['36800']  // 01.10.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Quarter 3"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.q3, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.06.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.07.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 30.09.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.10.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});
	QUnit.test('Test: "Date Filter - All Dates in the Period -> Quarter 4"', function (assert) {
		const testData = [
			['Dates'],
			['36799'], // 30.09.2000
			['36800'], // 01.10.2000
			['36891'], // 31.12.2000
			['36526']  // 01.01.2000
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Check data range
		checkFilterRef(assert, 0, 0, 4, 0);

		// Imitate choosing filter "Quarter 4"
		ws.autoFilters = createDynamicFilter(ws, Asc.c_oAscDynamicAutoFilter.q4, 0);

		//Checking work of filter
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 30.09.2000 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 01.10.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 31.12.2000 must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 01.01.2000 must be hidden');

		//Clearing data of sheet
		clearData(0, 0, 0, 4)
	});

	QUnit.test('Test: "Simple date filter apply"', function (assert) {
		const testData = [
			['Dates'],
			['20000'],
			['test1'],
			['50000'],
			['2/11/1930'],
			['test2'],
			['2/4/2237'],
			['3/4/2237'],
			['8/20/1994'],
			['6/16/1909']
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		let range = getRange(0, 0, 0, 0);
		ws.autoFilters.addAutoFilter(null, range);

		// Check data range
		checkFilterRef(assert, 0, 0, 9, 0);

		//apply filter
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.Filters);
		setFilterOptionsVisible(autoFiltersOptions, [{text: "20000", visible: false}, {
			text: "test1",
			visible: false
		}, {text: "2/11/1930", visible: false}]);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"1": 1, "2": 1, "4": 1});
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.Filters);
		setFilterOptionsVisible(autoFiltersOptions, [{text: "2/4/2237", visible: false}, {
			text: "3/4/2237",
			visible: false
		}]);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Checking work of filter
		checkHiddenRows(assert, testData, {"6": 1, "7": 1});

		//Clearing data of sheet
		clearData(0, 0, 0, 9)
	});

	QUnit.test('Test: "Date filter apply"', function (assert) {

		const testData = [
			['Dates'],
			['9/26/1902 2:24'],
			['9/26/1902 2:52'],
			['9/26/1902 3:07'],
			['6/22/1905 0:00'],
			['3/18/1908 0:00'],
			['12/13/1910 0:00'],
			['6/22/1905 2:24'],
			['6/22/1905 4:48'],
			['6/22/1905 7:12']
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		let range = getRange(0, 0, 0, 0);
		ws.autoFilters.addAutoFilter(null, range);

		// Check data range
		checkFilterRef(assert, 0, 0, 9, 0);

		//apply filter
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.Filters);
		//9/26/2002 2:24:00
		let aChangedVal = [{
			year: 1902,
			month: 9 - 1,
			day: 26,
			hour: 2,
			minute: 24,
			visible: false,
			dateTimeGrouping: Asc.EDateTimeGroup.datetimegroupHour
		}];
		aChangedVal.push({
			year: 1902,
			month: 9 - 1,
			day: 26,
			hour: 2,
			minute: 52,
			visible: true,
			dateTimeGrouping: Asc.EDateTimeGroup.datetimegroupMinute
		});
		aChangedVal.push({
			year: 1902,
			month: 9 - 1,
			day: 26,
			hour: 3,
			minute: 7,
			visible: true,
			dateTimeGrouping: Asc.EDateTimeGroup.datetimegroupHour
		});

		setFilterOptionsVisible(autoFiltersOptions, aChangedVal, true);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"1": 1}, " data filter apply ");
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		aChangedVal = [{
			val: 2000.3,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 7,
			minute: 12,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.2,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 4,
			minute: 48,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.1,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 2,
			minute: 24,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 1000.13,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 3,
			minute: 7,
			second: 12,
			dateTimeGrouping: 6
		}, {
			val: 1000.12,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 2,
			minute: 52,
			second: 48,
			dateTimeGrouping: 6
		}, {
			val: 1000.1,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 2,
			minute: 24,
			second: 0,
			dateTimeGrouping: 6
		}];

		setFilterOptionsVisible(autoFiltersOptions, aChangedVal, true);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"4": 1, "7": 1, "8": 1, "9": 1}, " data filter apply 2");
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		aChangedVal = [{
			val: 4000,
			visible: true,
			isDateFormat: true,
			year: 1910,
			month: 11,
			day: 13,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 3000,
			visible: true,
			isDateFormat: true,
			year: 1908,
			month: 2,
			day: 18,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.3,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 7,
			minute: 12,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.2,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 4,
			minute: 48,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.1,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 2,
			minute: 24,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 1000.13,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 3,
			minute: 7,
			second: 12,
			dateTimeGrouping: 2
		}, {
			val: 1000.12,
			visible: false,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 2,
			minute: 52,
			second: 48,
			dateTimeGrouping: 2
		}, {
			val: 1000.1,
			visible: false,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 2,
			minute: 24,
			second: 0,
			dateTimeGrouping: 2
		}];

		setFilterOptionsVisible(autoFiltersOptions, aChangedVal, true);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"1": 1, "2": 1, "4": 1, "7": 1, "8": 1 ,"9": 1}, " data filter apply 3");
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		aChangedVal = [{
			val: 4000,
			visible: true,
			isDateFormat: true,
			year: 1910,
			month: 11,
			day: 13,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 3000,
			visible: true,
			isDateFormat: true,
			year: 1908,
			month: 2,
			day: 18,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.3,
			visible: true,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 7,
			minute: 12,
			second: 0,
			dateTimeGrouping: 2
		}, {
			val: 2000.2,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 4,
			minute: 48,
			second: 0,
			dateTimeGrouping: 2
		}, {
			val: 2000.1,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 2,
			minute: 24,
			second: 0,
			dateTimeGrouping: 2
		}, {
			val: 2000,
			visible: false,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 2
		}, {
			val: 1000.13,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 3,
			minute: 7,
			second: 12,
			dateTimeGrouping: 6
		}, {
			val: 1000.12,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 2,
			minute: 52,
			second: 48,
			dateTimeGrouping: 6
		}, {
			val: 1000.1,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 2,
			minute: 24,
			second: 0,
			dateTimeGrouping: 6
		}];

		setFilterOptionsVisible(autoFiltersOptions, aChangedVal, true);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"4": 1, "7": 1, "8": 1}, " data filter apply 4");
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		aChangedVal = [{
			val: 4000,
			visible: true,
			isDateFormat: true,
			year: 1910,
			month: 11,
			day: 13,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 3000,
			visible: false,
			isDateFormat: true,
			year: 1908,
			month: 2,
			day: 18,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.3,
			visible: true,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 7,
			minute: 12,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.2,
			visible: true,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 4,
			minute: 48,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000.1,
			visible: true,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 2,
			minute: 24,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 2000,
			visible: true,
			isDateFormat: true,
			year: 1905,
			month: 5,
			day: 22,
			hour: 0,
			minute: 0,
			second: 0,
			dateTimeGrouping: 6
		}, {
			val: 1000.13,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 3,
			minute: 7,
			second: 12,
			dateTimeGrouping: 6
		}, {
			val: 1000.12,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 2,
			minute: 52,
			second: 48,
			dateTimeGrouping: 6
		}, {
			val: 1000.1,
			visible: true,
			isDateFormat: true,
			year: 1902,
			month: 8,
			day: 26,
			hour: 2,
			minute: 24,
			second: 0,
			dateTimeGrouping: 6
		}];

		setFilterOptionsVisible(autoFiltersOptions, aChangedVal, true);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"5": 1}, " data filter apply 4");
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//Clearing data of sheet
		clearData(0, 0, 0, 9)
	});

	QUnit.test('Test: "Custom date filter apply"', function (assert) {
		const testData = [
			['Dates'],
			['20000'],
			['test1'],
			['50000'],
			['2/11/1930'],
			['test2'],
			['2/4/2237'],
			['6/2/1906'],
			['8/20/1994'],
			['6/16/1909']
		];

		// Imitate filling rows with data, selection data range and add filter
		let range = getRange(0, 0, 0, 0);
		ws = getRangeWithData(ws, testData);
		ws.autoFilters.addAutoFilter(null, range);

		// Check data range
		checkFilterRef(assert, 0, 0, 9, 0);


		//***Before***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.isLessThan, val: "8/20/1994"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"2": 1, "3": 1, "5": 1, "6": 1, "8": 1}, " Before: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***After***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.isGreaterThan, val: "6500"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"2": 1, "5": 1, "7": 1, "9": 1}, " After: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***After or equal***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo, val: "6/16/1909"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"2": 1, "5": 1, "7": 1}, " After or equal: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***Before or equal***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo, val: "6/16/1909"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"8": 1
		}, " Before or equal: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***Between***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo, val: "6/16/1909"},
			{operator: Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo, val: "8/20/1994"}], true);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"2": 1, "3": 1, "5": 1, "6": 1, "7": 1}, " Between: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo, val: "6/16/1909"},
			{operator: Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo, val: "8/20/1994"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"2": 1, "5": 1}, " Between: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***equals***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.equals, val: "20000"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " equals1: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		//equals === only value with format
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.equals, val: "8/20/1994"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"9": 1
		}, " equals2: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.equals, val: "34566"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " equals3: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***doesNotEqual***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotEqual, val: "20000"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"1": 1}, " doesNotEqual1: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotEqual, val: "8/20/1994"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"8": 1}, " doesNotEqual2: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotEqual, val: "34566"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"8": 1}, " doesNotEqual3: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***begin with***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.beginsWith, val: "200"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " begin with1: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.beginsWith, val: "8/20/1994"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " begin with2: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.beginsWith, val: "tes"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"1": 1, "3": 1, "4": 1, "6": 1, "7": 1, "8": 1, "9": 1}, " begin with3: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***end with***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.endsWith, val: "000"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " end with1: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.endsWith, val: "20/1994"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " end with2: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.endsWith, val: "st2"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " end with2: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***does not begin with***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotBeginWith, val: "200"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {}, " does not begin with1: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotBeginWith, val: "8/20"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {}, " does not begin with2: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotBeginWith, val: "tes"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"2": 1, "5": 1}, " does not begin with3: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//contains
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.contains, val: "200"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " contains1: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.contains, val: "8/20"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {
			"1": 1,
			"2": 1,
			"3": 1,
			"4": 1,
			"5": 1,
			"6": 1,
			"7": 1,
			"8": 1,
			"9": 1
		}, " contains2: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.contains, val: "es"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"1": 1, "3": 1, "4": 1, "6": 1, "7": 1, "8": 1, "9": 1}, " contains3: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);


		//***does not contains***
		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotContain, val: "200"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {}, " does not contains1: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotContain, val: "8/20"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {}, " does not contains2: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		createCustomFilter(0, [{operator: Asc.c_oAscCustomAutoFilter.doesNotContain, val: "tes"}]);
		//Checking work of filter
		checkHiddenRows(assert, testData, {"2": 1, "5": 1}, " does not contains3: ");
		//clean filter
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		//Clearing data of sheet
		clearData(0, 0, 0, 9)
	});

	QUnit.test('Test: "Combine filter apply"', function (assert) {
		//TODO need check
		const testData = [
			['Dates'],
			['text1'],
			['text1'],
			['5/20/2020 0:00:00'],
			['6/20/2020 0:00:00'],
			['6/21/2020 0:00:00'],
			['6/22/2020 0:00:00'],
			['6/22/2020 10:00:00'],
			['6/22/2020 11:00:00'],
			['6/23/2020 11:10:00'],
			['6/24/2020 11:15:00'],
			['6/24/2020 11:20:10'],
			['6/24/2020 11:20:20'],
			['555']
		];

		// Imitate filling rows with data, selection data range and add filter
		ws = getRangeWithData(ws, testData);
		let range = getRange(0, 0, 0, 0);
		ws.autoFilters.addAutoFilter(null, range);

		// Check data range
		checkFilterRef(assert, 0, 0, 13, 0);

		//apply filter
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.Filters);
		//may exclude
		let aChangedVal = [{val:44006.472453703704,visible:true,isDateFormat:true,year:2020,month:5,day:24,hour:11,minute:20,second:20,dateTimeGrouping:4},{val:44006.472337962965,visible:true,isDateFormat:true,year:2020,month:5,day:24,hour:11,minute:20,second:10,dateTimeGrouping:4},{val:44006.46875,visible:true,isDateFormat:true,year:2020,month:5,day:24,hour:11,minute:15,second:0,dateTimeGrouping:4},{val:44005.46527777778,visible:true,isDateFormat:true,year:2020,month:5,day:23,hour:11,minute:10,second:0,dateTimeGrouping:4},{val:44004.458333333336,visible:true,isDateFormat:true,year:2020,month:5,day:22,hour:11,minute:0,second:0,dateTimeGrouping:4},{val:44004.416666666664,visible:true,isDateFormat:true,year:2020,month:5,day:22,hour:10,minute:0,second:0,dateTimeGrouping:4},{val:44004,visible:true,isDateFormat:true,year:2020,month:5,day:22,hour:0,minute:0,second:0,dateTimeGrouping:4},{val:44003,visible:true,isDateFormat:true,year:2020,month:5,day:21,hour:0,minute:0,second:0,dateTimeGrouping:4},{val:44002,visible:true,isDateFormat:true,year:2020,month:5,day:20,hour:0,minute:0,second:0,dateTimeGrouping:4},{val:43971,visible:false,isDateFormat:true,year:2020,month:4,day:20,hour:0,minute:0,second:0,dateTimeGrouping:4},{val:555,visible:true,isDateFormat:false},{val:"text1",visible:true,isDateFormat:false}]

		setFilterOptionsVisible(autoFiltersOptions, aChangedVal, true);
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);
		//Checking work of filter
		//checkHiddenRows(assert, testData, {"1": 1}, " combine filter apply 1");
		ws.autoFilters.isApplyAutoFilterInCell(range, true);

		clearData(0, 0, 0, 13);
	});

	QUnit.test('Test: "Open filter options"', function (assert) {
		const fOld_af_setDialogProp = wsView.af_setDialogProp;
		let oCurrentAnswer;
		wsView.af_setDialogProp = function (filterProp, tooltipPreview)
		{
			assert.deepEqual(filterProp, oCurrentAnswer, 'Check filter properties');
		}

		function select(activeR, activeC, r1, c1, r2, c2)
		{
			ws.selectionRange.ranges = [getRange(r1, c1, r2, c2)];
			ws.selectionRange.activeCell = new AscCommon.CellBase(activeR, activeC);
		}
		let testData = [
			["test1", "test2", "test3", "test4", "test5"],
			['1', '1', '1', '1', '1'],
			['1', '1', '1', '1', '1'],
			['1', '1', '1', '1', '1'],
			['1', '1', '1', '1', '1']
		];

		let range = ws.getRange4(1, 1);
		range.fillData(testData);
		ws.autoFilters.addAutoFilter(null, getRange(1, 1, 5, 1));

		select(2, 1, 2, 1, 2, 1);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		select(2, 2, 2, 2, 2, 2);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		select(2, 1, 1, 1, 2, 3);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		oCurrentAnswer = {id: null, colId: 0};
		select(1, 1, 1, 1, 2, 3);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		oCurrentAnswer = {id: null, colId: 1};
		select(1, 2, 1, 1, 2, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		oCurrentAnswer = {id: null, colId: 2};
		select(1, 3, 1, 1, 2, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');
		
		select(2, 3, 1, 1, 2, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		ws.getRange3(1, 0, 1, 1).merge();
		select(1, 1, 1, 0, 2, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');
		ws.getRange3(1, 0, 1, 1).unmerge();
		
		ws.getRange3(1, 4, 1, 5).merge();
		oCurrentAnswer = {id: null, colId: 3};
		select(1, 4, 1, 4, 1, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');
		ws.getRange3(1, 4, 1, 5).unmerge();

		ws.getRange3(1, 5, 1, 6).merge();
		oCurrentAnswer = {id: null, colId: 4};
		select(1, 6, 1, 5, 1, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');
		ws.getRange3(1, 5, 1, 6).unmerge();

		ws.getRange3(0, 0, 1, 1).merge();
		select(1, 1, 0, 0, 1, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');
		ws.getRange3(0, 0, 1, 1).unmerge();

		ws.AutoFilter.showButton(false);
		ws.AutoFilter.getFilterColumn(0).ShowButton = true;
		ws.AutoFilter.getFilterColumn(4).ShowButton = true;

		oCurrentAnswer = {id: null, colId: 1};
		select(1, 2, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		ws.getRange3(1, 2, 1, 3).merge();
		oCurrentAnswer = {id: null, colId: 1};
		select(1, 2, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');
		ws.getRange3(1, 2, 1, 3).unmerge();

		clearData(0, 0, 5, 5);

		range = ws.getRange4(1, 1);
		range.fillData(testData);
		ws.autoFilters.addAutoFilter('TableStyleLight9', getRange(1, 1, 5, 5));

		select(2, 1, 2, 1, 2, 1);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		select(2, 2, 2, 2, 2, 2);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		select(2, 1, 1, 1, 2, 3);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		oCurrentAnswer = {id: 0, colId: 0};
		select(1, 1, 1, 1, 2, 3);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		oCurrentAnswer = {id: 0, colId: 1};
		select(1, 2, 1, 1, 2, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		oCurrentAnswer = {id: 0, colId: 2};
		select(1, 3, 1, 1, 2, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		select(2, 3, 1, 1, 2, 5);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		ws.TableParts[0].AutoFilter.showButton(false);
		ws.TableParts[0].AutoFilter.getFilterColumn(0).ShowButton = true;
		ws.TableParts[0].AutoFilter.getFilterColumn(4).ShowButton = true;

		select(1, 2, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		select(1, 3, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		oCurrentAnswer = {id: 0, colId: 0};
		select(1, 1, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		ws.TableParts[0].AutoFilter.showButton(true);
		ws.autoFilters.changeFormatTableInfo(ws.TableParts[0].DisplayName, Asc.c_oAscChangeTableStyleInfo.rowHeader, false);

		select(1, 1, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		select(1, 2, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		select(2, 1, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		select(2, 2, 0, 0, 6, 6);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), false, 'Check non-opening filter options');

		ws.autoFilters.changeFormatTableInfo(ws.TableParts[0].DisplayName, Asc.c_oAscChangeTableStyleInfo.rowHeader, true);

		range = ws.getRange4(1, 8);
		range.fillData(testData);
		ws.autoFilters.addAutoFilter('TableStyleLight9', getRange(8, 1, 12, 5));

		oCurrentAnswer = {id: 1, colId: 0};
		select(1, 8, 1, 8, 2, 10);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		oCurrentAnswer = {id: 1, colId: 1};
		select(1, 9, 1, 8, 2, 12);
		assert.strictEqual(wsView.showAutoFilterOptionsFromActiveCell(), true, 'Check open filter options');

		clearData(1, 1, 5, 5);
		clearData(8, 1, 12, 5);

		wsView.af_setDialogProp = fOld_af_setDialogProp;
	});

	QUnit.module("Filters");
});
