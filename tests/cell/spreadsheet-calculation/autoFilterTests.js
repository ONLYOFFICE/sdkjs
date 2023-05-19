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

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);

	});
	QUnit.test('Test: "Date Filter - Today"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "Сегодня"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.today)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 15.05.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Yesterday"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "Вчера"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.yesterday)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 14.05.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Tomorrow"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "Завтра"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.tomorrow)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 16.05.2023 tomorrow must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - This week"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "На этой неделе"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.thisWeek)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 14.05.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 15.05.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 16.05.2023 tomorrow must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), false, 'Value 17.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Next week"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "На следующей неделе"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.nextWeek)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), false, 'Value 21.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Last week"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "На следующей неделе"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.lastWeek)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), false, 'Value 13.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), false, 'Value 07.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Last month"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "На прошлой неделе"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.lastMonth)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), false, 'Value 01.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - This month"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "В этом месяце"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.thisMonth)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), false, 'Value 13.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 14.05.2023 yesterday not must be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 15.05.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 16.05.2023 tomorrow must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), false, 'Value 17.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(6), false, 'Value 21.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(7), false, 'Value 28.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(8), false, 'Value 07.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), false, 'Value 01.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Next month"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "В следующем месяце"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.nextMonth)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), false, 'Value 01.06.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Next quarter"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "В следующем квартале"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.nextQuarter)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), false, 'Value 01.07.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - This quarter"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "В этом квартале"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.thisQuarter)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), false, 'Value 13.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 14.05.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 15.05.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 16.05.2023 tomorrow must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), false, 'Value 17.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(6), false, 'Value 21.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(7), false, 'Value 28.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(8), false, 'Value 07.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(9), false, 'Value 01.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(10), false, 'Value 01.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(11), false, 'Value 01.06.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Last quarter"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "В прошлом квартале"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.lastQuarter)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), false, 'Value 01.01.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - This year"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "В этом году"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.thisYear)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), false, 'Value 13.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 14.05.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 15.05.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 16.05.2023 tomorrow must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), false, 'Value 17.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(6), false, 'Value 21.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(7), false, 'Value 28.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(8), false, 'Value 07.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(9), false, 'Value 01.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(10), false, 'Value 01.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(11), false, 'Value 01.06.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(12), false, 'Value 01.07.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(13), false, 'Value 01.10.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(14), false, 'Value 01.01.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Next year"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "В cледующем году"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.nextYear)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), false, 'Value 01.01.2024 must not be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Last year"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562']  // 01.01.2022
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 17, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "В прошлом году"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.lastYear)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), true, 'Value 13.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(2), true, 'Value 14.05.2023 yesterday must be hidden');
		assert.strictEqual(ws.getRowHidden(3), true, 'Value 15.05.2023 today must be hidden');
		assert.strictEqual(ws.getRowHidden(4), true, 'Value 16.05.2023 tomorrow must be hidden');
		assert.strictEqual(ws.getRowHidden(5), true, 'Value 17.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), true, 'Value 07.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(9), true, 'Value 01.04.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(10), true, 'Value 01.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), true, 'Value 01.01.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must not be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), false, 'Value 01.01.2022 must not be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});
	QUnit.test('Test: "Date Filter - Year to date"', function (assert) {
		const testData = [
			['Dates'],
			['45059'], // 13.05.2023
			['45060'], // 14.05.2023
			['45061'], // 15.05.2023 today
			['45062'], // 16.05.2023
			['45063'], // 17.05.2023
			['45067'], // 18.05.2023
			['45074'], // 28.05.2023
			['45053'], // 07.05.2023
			['45017'], // 01.04.2023
			['45047'], // 01.05.2023
			['45078'], // 01.06.2023
			['45108'], // 01.07.2023
			['45200'], // 01.10.2023
			['44927'], // 01.01.2023
			['45292'], // 01.01.2024
			['45658'], // 01.01.2025
			['44562'], // 01.01.2022
			['45066']  // 20.05.2023
		];

		// Имитируем заполнение колонки данными выделяем диапозон и добавляем фильтр.
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		ws.autoFilters.addAutoFilter(null, getRange(0, 0, 0, 0));

		// Проверяем корректность диапозона данных
		assert.strictEqual(ws.AutoFilter.Ref.r1, 0, 'Check start point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c1, 0, 'Check start point column filter range');
		assert.strictEqual(ws.AutoFilter.Ref.r2, 18, 'Check finish point row filter range');
		assert.strictEqual(ws.AutoFilter.Ref.c2, 0, 'Check finish point column filter range');

		// Инициализируем настройки фильтра  и динамический фильтр
		let autoFiltersOptions = ws.autoFilters.getAutoFiltersOptions(ws, {colId: 0, id: null});
		let dynamicFilter = new Asc.DynamicFilter();

		// Имитируем выбор фильтра "C начала года"
		dynamicFilter.asc_setType(Asc.c_oAscDynamicAutoFilter.yearToDate)
		dynamicFilter.init(getRange(0,0,0,0));
		autoFiltersOptions.filter.asc_setType(c_oAscAutoFilterTypes.DynamicFilter);
		autoFiltersOptions.filter.asc_setFilter(dynamicFilter)
		ws.autoFilters.applyAutoFilter(autoFiltersOptions);

		//Проверяем корректность работы отображения строк
		assert.strictEqual(ws.getRowHidden(1), false, 'Value 13.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(2), false, 'Value 14.05.2023 yesterday must not be hidden');
		assert.strictEqual(ws.getRowHidden(3), false, 'Value 15.05.2023 today must not be hidden');
		assert.strictEqual(ws.getRowHidden(4), false, 'Value 16.05.2023 tomorrow must not be hidden');
		assert.strictEqual(ws.getRowHidden(5), false, 'Value 17.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(6), true, 'Value 21.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(7), true, 'Value 28.05.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(8), false, 'Value 07.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(9), false, 'Value 01.04.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(10), false, 'Value 01.05.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(11), true, 'Value 01.06.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(12), true, 'Value 01.07.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(13), true, 'Value 01.10.2023 must be hidden');
		assert.strictEqual(ws.getRowHidden(14), false, 'Value 01.01.2023 must not be hidden');
		assert.strictEqual(ws.getRowHidden(15), true, 'Value 01.01.2024 must be hidden');
		assert.strictEqual(ws.getRowHidden(16), true, 'Value 01.01.2025 must be hidden');
		assert.strictEqual(ws.getRowHidden(17), true, 'Value 01.01.2022 must be hidden');
		assert.strictEqual(ws.getRowHidden(18), true, 'Value 20.05.2022 must be hidden');

		//Очистка данных
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(0, 9, false);
		ws.removeCols(0, 1);
	});

	QUnit.module("CopyPaste");
});
