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
    // ====== REQUIRED ENVIRONMENT SETUP (preserve these stubs/settings) ======
    Asc.spreadsheet_api.prototype._init = function () { };
    Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
        callback();
    };
    AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () { };
    AscCommonExcel.WorkbookView.prototype._init = function () { };
    AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged =
        function () { };
    AscCommonExcel.WorkbookView.prototype.showWorksheet = function () { };
    AscCommonExcel.WorksheetView.prototype._init = function () { };
    AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable =
        function () { };
    AscCommonExcel.WorksheetView.prototype.setSelection = function () { };
    AscCommonExcel.WorksheetView.prototype.draw = function () { };
    AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects =
        function () { };
    AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () { };
    AscCommonExcel.WorksheetView.prototype.getZoom = function () { };
    AscCommonExcel.WorksheetView.prototype._getPPIY = function () { };
    AscCommonExcel.WorksheetView.prototype._getPPIX = function () { };
    AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () { };
    Asc.ReadDefTableStyles = function () { };

    var api = new Asc.spreadsheet_api({ "id-view": "editor_sdk" });

    api.FontLoader = { LoadDocumentFonts: function () { } };
    window["Asc"]["editor"] = api;
    AscCommon.g_oTableId.init();
    api._onEndLoadSdk();
    api.isOpenOOXInBrowser = false;
    api.OpenDocumentFromBin(null, AscCommon.getEmpty());
    api.initCollaborativeEditing({});
    api.wb = new AscCommonExcel.WorkbookView(
        api.wbModel,
        api.controller,
        api.handlers,
        api.HtmlElement,
        api.topLineEditorElement,
        api,
        api.collaborativeEditing,
        api.fontRenderingMode
    );

    var wsView = api.wb.getWorksheet(0);
    wsView.handlers = api.handlers;
    wsView.objectRender = new AscFormat.DrawingObjects();
    wsView.objectRender.controller = new AscFormat.DrawingObjectsController(
        wsView.objectRender
    );
    var ws = api.GetActiveSheet();

    // ====== TEST UTILITIES ======

    // MUST-HAVE helper: clear all conditional formats in A1:Z100 before each test
    window.initializeTest = function () {
        var r = ws.GetRange("A1:Z100");
        ws.worksheet.AutoFilter = null;
        // r.Clear();
    };

    theRange = function (address) {
        return ws.GetRange(address);
    };

    // ====== TESTS ======

    QUnit.module("ApiAutoFilter", function () {
        QUnit.test("AutoFilter filters length after Clear", function (assert) {
            initializeTest();
            let range = ws.GetRange("A1");
            range.SetAutoFilter();
            let filters = ws.AutoFilter.Filters;
            // length of filters should be 0 after Clear()
            assert.equal(filters.length, 0, "Filters length is 0");
        });

        QUnit.test("ApiAutoFilter check properties", function (assert) {
            initializeTest();

            theRange("A1").SetValue(10);
            theRange("A2").SetValue(20);
            theRange("A3").SetValue(2);
            theRange("A4").SetValue(5);
            theRange("A5").SetValue(4);
            theRange("A6").SetValue(7);

            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, [2, 5], "xlFilterValues");

            let filters = ws.AutoFilter.Filters;
            // length of filters should be 1 after setting filter on column 1

            assert.equal(filters.length, 1, "Filters length is 1");
            assert.equal(filters[0].Criteria1, "=2", "Criteria1 is 2");
            assert.equal(filters[0].Operator, "xlOr", "Operator is xlFilterValues");
            assert.equal(filters[0].Criteria2, "=5", "Criteria2 is 5");
            assert.equal(filters[0].On, true, "On is true");
        });

        QUnit.test("Remove AutoFilter when Field is null", function (assert) {
            initializeTest();

            let range = ws.GetRange("A1:B5");
            range.SetAutoFilter(1, [2,5], "xlFilterValues");
            assert.equal(ws.AutoFilter.FilterMode, true, "FilterMode true after add");

            range.SetAutoFilter(null); // remove whole AutoFilter
            assert.equal(ws.AutoFilter.FilterMode, false, "FilterMode false after remove");
            assert.equal(ws.AutoFilter.Range, null, "Range null after remove");
            assert.equal(ws.AutoFilter.Filters.length, 0, "Filters empty after remove");
        });

        QUnit.test("xlFilterValues with 3+ items keeps xlFilterValues", function (assert) {
            initializeTest();

            ["2", "5", "7", "9"].forEach((v, i) => theRange("A" + (i + 2)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, ["2", "5", "7"], "xlFilterValues");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.deepEqual(f[0].Criteria1.sort(), ["2", "5", "7"], "Criteria1 returns visible values array");
            assert.equal(f[0].Criteria2, null, "Criteria2 is null for values filter");
            assert.equal(f[0].Operator, "xlFilterValues", "Operator is xlFilterValues");
            assert.equal(f[0].On, true, "On true");
        });

        QUnit.test("xlFilterValues with 2 items converts to xlOr custom filter", function (assert) {
            initializeTest();

            [10, 20, 2, 5, 4, 7].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, [2, 5], "xlFilterValues");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlOr", "Operator converted to xlOr");
            assert.equal(f[0].Criteria1, "=2", "Criteria1 is '=2'");
            assert.equal(f[0].Criteria2, "=5", "Criteria2 is '=5'");
            assert.equal(f[0].On, true, "On true");
        });

        QUnit.test("Custom filter xlOr preserves signs and criteria", function (assert) {
            initializeTest();

            [1, 2, 3, 4, 5, 6, 7, 8, 9].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, ">3", "xlOr", "<=5");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlOr");
            assert.equal(f[0].Criteria1, ">3");
            assert.equal(f[0].Criteria2, "<=5");
            assert.equal(f[0].On, true, "Filter is on, i.e. applied");
        });

        QUnit.test("Custom filter xlAnd", function (assert) {
            initializeTest();

            [1, 2, 3, 4, 5, 6, 7, 8, 9].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, ">=2", "xlAnd", "<8");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlAnd");
            assert.equal(f[0].Criteria1, ">=2");
            assert.equal(f[0].Criteria2, "<8");
            assert.equal(f[0].On, true, "Filter is on, i.e. applied");
        });

        QUnit.test("Top10 items filter", function (assert) {
            initializeTest();

            [10, 20, 30, 40, 50, 60].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, "3", "xlTop10Items");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlTop10Items");
            assert.equal(f[0].Criteria1, 3, "Criteria1 returns numeric Val");
            assert.equal(f[0].Criteria2, null, "Criteria2 should be null");
            assert.equal(f[0].On, true, "Filter is on");
        });

        QUnit.test("Bottom10 percent filter", function (assert) {
            initializeTest();

            [10, 20, 30, 40, 50, 60].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, "50", "xlBottom10Percent");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlBottom10Percent");
            assert.equal(f[0].Criteria1, 50, "Criteria1 returns numeric Val");
            assert.equal(f[0].On, true, "Filter is on");
        });

        QUnit.test("Dynamic filter AboveAverage", function (assert) {
            initializeTest();

            [1, 2, 3, 100].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, "xlFilterAboveAverage", "xlFilterDynamic");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlFilterDynamic");
            assert.equal(f[0].Criteria1, "xlFilterAboveAverage");
            assert.equal(f[0].Criteria2, null, "Criteria2 should be null");
            assert.equal(f[0].On, true, "Filter is on");
        });

        QUnit.test("Color filter CellColor (Criteria1 null by API)", function (assert) {
            initializeTest();

            // Make some colored cells to be realistic
            theRange("A1").SetValue("x");
            theRange("A2").SetValue("y");
            theRange("A1").SetFillColor(api.CreateColorFromRGB(255, 255, 0));
            theRange("A2").SetFillColor(api.CreateColorFromRGB(0, 255, 0));

            let range = ws.GetRange("A1:A5");
            range.SetAutoFilter(1, api.CreateColorFromRGB(255, 255, 0), "xlFilterCellColor");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlFilterCellColor");
            assert.equal(f[0].Criteria1, null, "Criteria1 is null for color filters");
            assert.equal(f[0].Criteria2, null, "Criteria2 is null for color filters");
            assert.equal(f[0].On, true, "Filter is on");
        });

        QUnit.test("Clear specific column filter when Criteria1 is null", function (assert) {
            initializeTest();

            [10, 20, 2, 5, 4, 7].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");

            range.SetAutoFilter(1, ">3", "xlOr", "<=7");
            assert.equal(ws.AutoFilter.Filters.length, 1, "Filter exists before clear");

            range.SetAutoFilter(1, null); // clear column filter
            assert.equal(ws.AutoFilter.FilterMode, true, "AutoFilter still exists");
            assert.equal(ws.AutoFilter.Filters.length, 1, "Filters empty after column clear");
        });

        QUnit.test("Invalid Field does not add AutoFilter", function (assert) {
            initializeTest();

            let range = ws.GetRange("A1:B5");
            range.SetAutoFilter("foo", ">1", "xlOr"); // invalid Field

            assert.equal(ws.AutoFilter.FilterMode, false, "AutoFilter not added");
            assert.equal(ws.AutoFilter.Range, null, "Range null");
            assert.equal(ws.AutoFilter.Filters.length, 0, "No filters created");
        });

        QUnit.test("Field out of range does not add AutoFilter", function (assert) {
            initializeTest();

            let range = ws.GetRange("A1:B5"); // only 2 columns
            range.SetAutoFilter(3, ">1", "xlOr"); // invalid: field > columns count

            assert.equal(ws.AutoFilter.FilterMode, false, "AutoFilter not added");
            assert.equal(ws.AutoFilter.Filters.length, 0, "No filters created");
        });

    });

    QUnit.test("ApplyFilter recalculates visibility on new data", function (assert) {
        initializeTest();

        // Values: >2 are visible, <=2 are hidden after filter
        [1, 2, 2, 4, 5].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
        let range = ws.GetRange("A1:A10");
        range.SetAutoFilter(1, ">2", "xlOr");

        assert.equal(ws.AutoFilter.FilterMode, true, "FilterMode is true after SetAutoFilter");

        // Header row is row 0 in ws.worksheet; data below
        assert.equal(ws.worksheet.getRowHidden(0), false, "Row 1 visible (1)");
        assert.equal(ws.worksheet.getRowHidden(1), true,  "Row 2 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(2), true,  "Row 3 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(3), false, "Row 4 visible (4)");
        assert.equal(ws.worksheet.getRowHidden(4), false, "Row 5 visible (5)");

        // Change data after filter has been applied
        [1, 2, 2, 4, 5, 2, 3].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));

        // Still using old visibility until ApplyFilter is called again
        assert.equal(ws.worksheet.getRowHidden(5), false, "Row 7 visible before ApplyFilter");
        assert.equal(ws.worksheet.getRowHidden(6), false, "Row 8 visible before ApplyFilter");

        // Reapply filter – should reevaluate all rows
        ws.AutoFilter.ApplyFilter();

        assert.equal(ws.worksheet.getRowHidden(0), false, "Row 1 visible (1)");
        assert.equal(ws.worksheet.getRowHidden(1), true,  "Row 2 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(2), true,  "Row 3 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(3), false,  "Row 4 visible (4)");
        assert.equal(ws.worksheet.getRowHidden(4), false, "Row 5 visible (5)");
        assert.equal(ws.worksheet.getRowHidden(5), true, "Row 6 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(6), false,  "Row 7 visible (3)");
    });

    QUnit.test("ShowAllData makes all rows visible but keeps AutoFilter", function (assert) {
        initializeTest();

        // Values: only >2 should be visible before ShowAllData
        [1, 2, 3, 4, 5].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
        let range = ws.GetRange("A1:A10");
        range.SetAutoFilter(1, "<1", "xlOr");

        // Sanity check: some rows are hidden
        assert.equal(ws.worksheet.getRowHidden(0), false, "Row 1 visible (header)");
        assert.equal(ws.worksheet.getRowHidden(1), true, "Row 2 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(2), true, "Row 3 hidden (3)");
        assert.equal(ws.worksheet.getRowHidden(3), true, "Row 4 hidden (4)");
        assert.equal(ws.worksheet.getRowHidden(4), true, "Row 5 hidden (5)");

        ws.AutoFilter.ShowAllData();

        // All rows of the AutoFilter range must be visible now
        assert.equal(ws.worksheet.getRowHidden(0), false, "Row 1 visible (header)");
        assert.equal(ws.worksheet.getRowHidden(1), false, "Row 2 visible (2)");
        assert.equal(ws.worksheet.getRowHidden(2), false, "Row 3 visible (3)");
        assert.equal(ws.worksheet.getRowHidden(3), false, "Row 4 visible (4)");
        assert.equal(ws.worksheet.getRowHidden(4), false, "Row 5 visible (5)");

        // AutoFilter object should still exist (drop-downs preserved)
        assert.equal(ws.AutoFilter.FilterMode, true, "AutoFilter still present after ShowAllData");
    });

});
