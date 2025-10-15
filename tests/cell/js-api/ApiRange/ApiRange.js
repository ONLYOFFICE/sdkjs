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
    Asc.spreadsheet_api.prototype._init = function () {};
    Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
        callback();
    };
    AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {};
    AscCommonExcel.WorkbookView.prototype._init = function () {};
    AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged =
        function () {};
    AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {};
    AscCommonExcel.WorksheetView.prototype._init = function () {};
    AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable =
        function () {};
    AscCommonExcel.WorksheetView.prototype.setSelection = function () {};
    AscCommonExcel.WorksheetView.prototype.draw = function () {};
    AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects =
        function () {};
    AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () {};
    AscCommonExcel.WorksheetView.prototype.getZoom = function () {};
    AscCommonExcel.WorksheetView.prototype._getPPIY = function () {};
    AscCommonExcel.WorksheetView.prototype._getPPIX = function () {};
    AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {};
    Asc.ReadDefTableStyles = function () {};

    var api = new Asc.spreadsheet_api({
        "id-view": "editor_sdk",
    });

    api.FontLoader = {
        LoadDocumentFonts: function () {},
    };
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

    api.AddDefName('super', 'Sheet1!$A$1:$F$4');
    api.AddDefName('negativeIndexColumn', 'Sheet1!$F$5:$I$14');
    api.AddDefName('negativeIndexRow', 'Sheet1!$F$8:$H$13');
    api.AddDefName('outOfRangeColumn', 'Sheet1!$H$5:$I$14');
    api.AddDefName('outOfRangeRow', 'Sheet1!$F$11:$H$13');

    // QUnit.module("Test api range sort for all editors");

    QUnit.test("SetSort: single cell, header yes, orientation column, ascending order", function(assert) {
        ws.GetRange("A1").SetValue("Header");
        ws.GetRange("A1:A1").SetSort("A1", "xlAscending", null, null, null, null, "xlYes", "xlSortColumns");
        // Header should remain unchanged
        assert.strictEqual(ws.GetRange("A1").GetValue(), "Header", "Single cell header remains unchanged");
    });

    QUnit.test( "Multi-key sorting by A then C (A asc, C desc) with key2=null (row)", function (assert) {
        // Transposed data: each value set in a column, rows are keys
        // r1: (2,1,2,1,2,1)
        // r2: (b,d,a,c,z,a)
        // r3: (3,1,5,2,2,4)
        ws.GetRange("A1").SetValue("2"); ws.GetRange("A2").SetValue("b"); ws.GetRange("A3").SetValue("3");
        ws.GetRange("B1").SetValue("1"); ws.GetRange("B2").SetValue("d"); ws.GetRange("B3").SetValue("1");
        ws.GetRange("C1").SetValue("2"); ws.GetRange("C2").SetValue("a"); ws.GetRange("C3").SetValue("5");
        ws.GetRange("D1").SetValue("1"); ws.GetRange("D2").SetValue("c"); ws.GetRange("D3").SetValue("2");
        ws.GetRange("E1").SetValue("2"); ws.GetRange("E2").SetValue("z"); ws.GetRange("E3").SetValue("2");
        ws.GetRange("F1").SetValue("1"); ws.GetRange("F2").SetValue("a"); ws.GetRange("F3").SetValue("4");

        // Sort by row 1 ascending, skip key2 (null), then by row 3 descending
        ws.GetRange("A1:F3").SetSort("A1", "xlAscending", null, null, "A3", "xlDescending", "xlNo", "xlSortRows");

        // Expected order (row 1 asc, then row 3 desc; row 2 ignored as tie-breaker):
        // r1: (1,1,1,2,2,2)
        // r2: (a,c,d,a,b,z)
        // r3: (4,2,1,5,3,2)
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1", "Row1 A1");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "1", "Row1 B1");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "1", "Row1 C1");
        assert.strictEqual(ws.GetRange("D1").GetValue(), "2", "Row1 D1");
        assert.strictEqual(ws.GetRange("E1").GetValue(), "2", "Row1 E1");
        assert.strictEqual(ws.GetRange("F1").GetValue(), "2", "Row1 F1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "a", "Row2 A1");
        assert.strictEqual(ws.GetRange("B2").GetValue(), "c", "Row2 B1");
        assert.strictEqual(ws.GetRange("C2").GetValue(), "d", "Row2 C1");
        assert.strictEqual(ws.GetRange("D2").GetValue(), "a", "Row2 D1");
        assert.strictEqual(ws.GetRange("E2").GetValue(), "b", "Row2 E1");
        assert.strictEqual(ws.GetRange("F2").GetValue(), "z", "Row2 F1");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "4", "Row3 A1");
        assert.strictEqual(ws.GetRange("B3").GetValue(), "2", "Row3 B1");
        assert.strictEqual(ws.GetRange("C3").GetValue(), "1", "Row3 C1");
        assert.strictEqual(ws.GetRange("D3").GetValue(), "5", "Row3 D1");
        assert.strictEqual(ws.GetRange("E3").GetValue(), "3", "Row3 E1");
        assert.strictEqual(ws.GetRange("F3").GetValue(), "2", "Row3 F1");
    });

    QUnit.test("SetSort: handles empty cells (row)", function (assert) {
        ws.GetRange("A1").SetValue("");
        ws.GetRange("B1").SetValue("2");
        ws.GetRange("C1").SetValue("");
        ws.GetRange("D1").SetValue("1");
        ws.GetRange("A1:D1").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortRows");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "2");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "");
        assert.strictEqual(ws.GetRange("D1").GetValue(), "");
    });

    QUnit.test("SetSort: handles duplicate values (row)", function (assert) {
        ws.GetRange("A1").SetValue("2");
        ws.GetRange("B1").SetValue("2");
        ws.GetRange("C1").SetValue("1");
        ws.GetRange("D1").SetValue("1");
        ws.GetRange("A1:D1").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortRows");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "2");
        assert.strictEqual(ws.GetRange("D1").GetValue(), "2");
    });

    QUnit.test("SetSort: handles mixed types (row)", function (assert) {
        ws.GetRange("A1").SetValue("2");
        ws.GetRange("B1").SetValue("apple");
        ws.GetRange("C1").SetValue("1");
        ws.GetRange("D1").SetValue("banana");
        ws.GetRange("A1:D1").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortRows");
        // Numbers before strings
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "2");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "apple");
        assert.strictEqual(ws.GetRange("D1").GetValue(), "banana");
    });

    QUnit.test("SetSort: handles single column (row)", function (assert) {
        ws.GetRange("A1").SetValue("5");
        ws.GetRange("A1:A1").SetSort("A1", "xlDescending", null, null, null, null, "xlNo", "xlSortRows");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "5");
    });

    QUnit.test("SetSort: handles all identical values (row)", function (assert) {
        ws.GetRange("A1").SetValue("x");
        ws.GetRange("B1").SetValue("x");
        ws.GetRange("C1").SetValue("x");
        ws.GetRange("D1").SetValue("x");
        ws.GetRange("A1:D1").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortRows");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "x");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "x");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "x");
        assert.strictEqual(ws.GetRange("D1").GetValue(), "x");
    });

    QUnit.test("SetSort: non-existent key row", function (assert) {
        ws.GetRange("A1").SetValue("1");
        ws.GetRange("B1").SetValue("2");
        ws.GetRange("C1").SetValue("3");
        ws.GetRange("A1:C1").SetSort("A5", "xlAscending", null, null, null, null, "xlNo", "xlSortRows");
        // Should not crash, order unchanged
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "2");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "3");
    });

    QUnit.test("SetSort: with header column (row)", function (assert) {
        ws.GetRange("A1").SetValue("Header");
        ws.GetRange("B1").SetValue("3");
        ws.GetRange("C1").SetValue("1");
        ws.GetRange("D1").SetValue("2");
        ws.GetRange("A1:D1").SetSort("B1", "xlAscending", null, null, null, null, "xlYes", "xlSortRows");
        // Header should remain at left
        assert.strictEqual(ws.GetRange("A1").GetValue(), "Header");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "2");
        assert.strictEqual(ws.GetRange("D1").GetValue(), "3");
    });

    QUnit.test("SetSort: handles empty cells", function (assert) {
        ws.GetRange("A1").SetValue("");
        ws.GetRange("A2").SetValue("2");
        ws.GetRange("A3").SetValue("");
        ws.GetRange("A4").SetValue("1");
        ws.GetRange("A1:A4").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "2");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "");
    });

    QUnit.test("SetSort: handles duplicate values", function (assert) {
        ws.GetRange("A1").SetValue("2");
        ws.GetRange("A2").SetValue("2");
        ws.GetRange("A3").SetValue("1");
        ws.GetRange("A4").SetValue("1");
        ws.GetRange("A1:A4").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "1");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "2");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "2");
    });

    QUnit.test("SetSort: handles mixed types", function (assert) {
        ws.GetRange("A1").SetValue("2");
        ws.GetRange("A2").SetValue("apple");
        ws.GetRange("A3").SetValue("1");
        ws.GetRange("A4").SetValue("banana");
        ws.GetRange("A1:A4").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");

        // Numbers before strings
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "2");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "apple");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "banana");
    });

    QUnit.test("SetSort: handles single row", function (assert) {
        ws.GetRange("A1").SetValue("5");
        ws.GetRange("A1:A1").SetSort("A1", "xlDescending", null, null, null, null, "xlNo", "xlSortColumns");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "5");
    });

    QUnit.test("SetSort: handles all identical values", function (assert) {
        ws.GetRange("A1").SetValue("x");
        ws.GetRange("A2").SetValue("x");
        ws.GetRange("A3").SetValue("x");
        ws.GetRange("A4").SetValue("x");
        ws.GetRange("A1:A4").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "x");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "x");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "x");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "x");
    });

    QUnit.test("SetSort: non-existent key column", function (assert) {
        ws.GetRange("A1").SetValue("1");
        ws.GetRange("A2").SetValue("2");
        ws.GetRange("A3").SetValue("3");
        ws.GetRange("A1:A3").SetSort("Z1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");
        // Should not crash, order unchanged
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "2");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "3");
    });

    QUnit.test("SetSort: with header row", function (assert) {
        ws.GetRange("A1").SetValue("Header");
        ws.GetRange("A2").SetValue("3");
        ws.GetRange("A3").SetValue("1");
        ws.GetRange("A4").SetValue("2");
        ws.GetRange("A1:A4").SetSort("A2", "xlAscending", null, null, null, null, "xlYes", "xlSortColumns");
        // Header should remain at top
        assert.strictEqual(ws.GetRange("A1").GetValue(), "Header");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "1");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "2");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "3");
    });

    QUnit.test("SetSort: orientation row (sort by row)", function (assert) {
        ws.GetRange("A1").SetValue("1");
        ws.GetRange("B1").SetValue("3");
        ws.GetRange("C1").SetValue("2");
        ws.GetRange("A1:C1").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortRows");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "2");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "3");
    });

    QUnit.test("Test asc_sortRanges", function (assert) {
        console.log("here");
        ws.GetRange("A1").SetValue("1");
        ws.GetRange("A2").SetValue("2");
        ws.GetRange("A3").SetValue("3");
        ws.GetRange("A4").SetValue("4");
        ws.GetRange("B1").SetValue("4");
        ws.GetRange("B2").SetValue("3");
        ws.GetRange("B3").SetValue("2");
        ws.GetRange("B4").SetValue("1");
        ws.GetRange("C1").SetValue("3");
        ws.GetRange("C2").SetValue("4");
        ws.GetRange("C3").SetValue("1");
        ws.GetRange("C4").SetValue("2");
        // ApiRange
        ws.GetRange("A1:C4").SetSort("A1", "xlDescending", null, null, null, null, "xlNo", "xlSortColumns");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "4", "Check A1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "3", "Check A2");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "2", "Check A3");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "1", "Check A4");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "1", "Check B1");
        assert.strictEqual(ws.GetRange("B2").GetValue(), "2", "Check B2");
        assert.strictEqual(ws.GetRange("B3").GetValue(), "3", "Check B3");
        assert.strictEqual(ws.GetRange("B4").GetValue(), "4", "Check B4");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "2", "Check C1");
        assert.strictEqual(ws.GetRange("C2").GetValue(), "1", "Check C2");
        assert.strictEqual(ws.GetRange("C3").GetValue(), "4", "Check C3");
        assert.strictEqual(ws.GetRange("C4").GetValue(), "3", "Check C4");
    });

    QUnit.test("One column range sorting (A1:A5) asc/desc", function (assert) {
        // Set up
        ws.GetRange("A1").SetValue("5");
        ws.GetRange("A2").SetValue("3");
        ws.GetRange("A3").SetValue("4");
        ws.GetRange("A4").SetValue("1");
        ws.GetRange("A5").SetValue("2");

        // Ascending by A
        ws.GetRange("A1:A5").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1", "Asc A1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "2", "Asc A2");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "3", "Asc A3");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "4", "Asc A4");
        assert.strictEqual(ws.GetRange("A5").GetValue(), "5", "Asc A5");

        // Descending by A (same range)
        ws.GetRange("A1:A5").SetSort("A1", "xlDescending", null, null, null, null, "xlNo", "xlSortColumns");

        assert.strictEqual(ws.GetRange("A1").GetValue(), "5", "Desc A1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "4", "Desc A2");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "3", "Desc A3");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "2", "Desc A4");
        assert.strictEqual(ws.GetRange("A5").GetValue(), "1", "Desc A5");

        // Ascending by A
        ws.GetRange("B1").SetSort("A1", "xlDescending", null, null, null, null, "xlNo", "xlSortColumns");

        assert.strictEqual(ws.GetRange("A1").GetValue(), "5", "Asc A1");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "4", "Asc A2");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "3", "Asc A3");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "2", "Asc A4");
        assert.strictEqual(ws.GetRange("A5").GetValue(), "1", "Asc A5");
    });

    QUnit.test("Two-column range sorting by first key (A)", function (assert) {
        // Reset grid A1:B4
        // Rows: [A,B] -> mix ensures row movement is visible
        ws.GetRange("A1").SetValue("3");
        ws.GetRange("B1").SetValue("a");
        ws.GetRange("A2").SetValue("1");
        ws.GetRange("B2").SetValue("c");
        ws.GetRange("A3").SetValue("2");
        ws.GetRange("B3").SetValue("b");
        ws.GetRange("A4").SetValue("4");
        ws.GetRange("B4").SetValue("d");

        // Sort A1:B4 by column A ascending
        ws.GetRange("A1:B4").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");
        // Expect rows ordered by A: (1,c), (2,b), (3,a), (4,d)
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1", "Row1 A");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "c", "Row1 B");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "2", "Row2 A");
        assert.strictEqual(ws.GetRange("B2").GetValue(), "b", "Row2 B");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "3", "Row3 A");
        assert.strictEqual(ws.GetRange("B3").GetValue(), "a", "Row3 B");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "4", "Row4 A");
        assert.strictEqual(ws.GetRange("B4").GetValue(), "d", "Row4 B");
    });

    // 3) Two-column range sorting on key = second column (B)
    QUnit.test("Two-column range sorting by second key (B)", function (assert) {
        // Reset grid A1:B4 (same data as prior test)
        ws.GetRange("A1").SetValue("3");
        ws.GetRange("B1").SetValue("a");
        ws.GetRange("A2").SetValue("1");
        ws.GetRange("B2").SetValue("c");
        ws.GetRange("A3").SetValue("2");
        ws.GetRange("B3").SetValue("b");
        ws.GetRange("A4").SetValue("4");
        ws.GetRange("B4").SetValue("d");

        // Sort A1:B4 by column B ascending (key 'B1')
        ws.GetRange("A1:B4").SetSort("B1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");

        // Expect rows ordered by B: a,b,c,d with A moved accordingly -> (3,a),(2,b),(1,c),(4,d)
        assert.strictEqual(ws.GetRange("A1").GetValue(), "3", "Row1 A (a)");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "a", "Row1 B (a)");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "2", "Row2 A (b)");
        assert.strictEqual(ws.GetRange("B2").GetValue(), "b", "Row2 B (b)");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "1", "Row3 A (c)");
        assert.strictEqual(ws.GetRange("B3").GetValue(), "c", "Row3 B (c)");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "4", "Row4 A (d)");
        assert.strictEqual(ws.GetRange("B4").GetValue(), "d", "Row4 B (d)");
    });

    // 4) Sorting with "invalid" sort range should not sort anything
    // (Practical, side-effect-free definition: a single-cell range is not sortable)
    QUnit.test("Invalid sort range (single cell) is a no-op", function (assert) {
        ws.GetRange("A1").SetValue("42");
        // Sorting a single cell can't reorder anything
        ws.GetRange("A1:A1").SetSort("A1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");
        assert.strictEqual(ws.GetRange("A1").GetValue(), "42", "Single cell remains unchanged");
    });

    // 5) Sorting works only when sorting key and range intersect
    QUnit.test( "No sort when key does not intersect sort range", function (assert) {
        // Prepare A1:B4 data that would visibly change if any sort occurred
        ws.GetRange("A1").SetValue("5");
        ws.GetRange("B1").SetValue("w");
        ws.GetRange("A2").SetValue("3");
        ws.GetRange("B2").SetValue("y");
        ws.GetRange("A3").SetValue("4");
        ws.GetRange("B3").SetValue("x");
        ws.GetRange("A4").SetValue("1");
        ws.GetRange("B4").SetValue("z");

        // Place an unrelated key in column C so it does NOT intersect A1:B4
        ws.GetRange("C1").SetValue("10");
        ws.GetRange("C2").SetValue("20");
        ws.GetRange("C3").SetValue("30");
        ws.GetRange("C4").SetValue("40");

        // Attempt to sort A1:B4 using key at C1 (non-intersecting)
        ws.GetRange("A1:B4").SetSort("C1", "xlAscending", null, null, null, null, "xlNo", "xlSortColumns");

        // Expect NO changes in A/B
        assert.strictEqual(ws.GetRange("A1").GetValue(), "5", "A1 unchanged");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "w", "B1 unchanged");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "3", "A2 unchanged");
        assert.strictEqual(ws.GetRange("B2").GetValue(), "y", "B2 unchanged");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "4", "A3 unchanged");
        assert.strictEqual(ws.GetRange("B3").GetValue(), "x", "B3 unchanged");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "1", "A4 unchanged");
        assert.strictEqual(ws.GetRange("B4").GetValue(), "z", "B4 unchanged");
    });

    QUnit.test("Multi-key sorting by A then B (A asc, B desc)", function (assert) {
        // Data with duplicates in A so B can break ties
        // Original rows: (A,B)
        // 1: (2,b)
        // 2: (1,d)
        // 3: (2,a)
        // 4: (1,c)
        // 5: (3,e)
        ws.GetRange("A1").SetValue("2");
        ws.GetRange("B1").SetValue("b");
        ws.GetRange("A2").SetValue("1");
        ws.GetRange("B2").SetValue("d");
        ws.GetRange("A3").SetValue("2");
        ws.GetRange("B3").SetValue("a");
        ws.GetRange("A4").SetValue("1");
        ws.GetRange("B4").SetValue("c");
        ws.GetRange("A5").SetValue("3");
        ws.GetRange("B5").SetValue("e");

        // Sort A1:B5 by A ascending, then B descending
        ws.GetRange("A1:B5").SetSort(
            "A1",
            "xlAscending",
            "B1",
            "xlDescending",
            null,
            null,
            "xlNo",
            "xlSortColumns"
        );

        // Expected: (1,d), (1,c), (2,b), (2,a), (3,e)
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1", "Row1 A");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "d", "Row1 B");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "1", "Row2 A");
        assert.strictEqual(ws.GetRange("B2").GetValue(), "c", "Row2 B");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "2", "Row3 A");
        assert.strictEqual(ws.GetRange("B3").GetValue(), "b", "Row3 B");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "2", "Row4 A");
        assert.strictEqual(ws.GetRange("B4").GetValue(), "a", "Row4 B");
        assert.strictEqual(ws.GetRange("A5").GetValue(), "3", "Row5 A");
        assert.strictEqual(ws.GetRange("B5").GetValue(), "e", "Row5 B");
    });

    QUnit.test("Multi-key sorting by A then B then C (A asc, B asc, C desc)", function (assert) {
        // Initial data (A,B,C) with ties in A and B so C can break ties:
        // r1: (2,b,3)
        // r2: (1,d,1)
        // r3: (2,a,5)
        // r4: (1,c,2)
        // r5: (3,e,9)
        // r6: (2,a,1)
        // r7: (2,b,2)
        // r8: (1,c,1)
        ws.GetRange("A1").SetValue("2");
        ws.GetRange("B1").SetValue("b");
        ws.GetRange("C1").SetValue("3");
        ws.GetRange("A2").SetValue("1");
        ws.GetRange("B2").SetValue("d");
        ws.GetRange("C2").SetValue("1");
        ws.GetRange("A3").SetValue("2");
        ws.GetRange("B3").SetValue("a");
        ws.GetRange("C3").SetValue("5");
        ws.GetRange("A4").SetValue("1");
        ws.GetRange("B4").SetValue("c");
        ws.GetRange("C4").SetValue("2");
        ws.GetRange("A5").SetValue("3");
        ws.GetRange("B5").SetValue("e");
        ws.GetRange("C5").SetValue("9");
        ws.GetRange("A6").SetValue("2");
        ws.GetRange("B6").SetValue("a");
        ws.GetRange("C6").SetValue("1");
        ws.GetRange("A7").SetValue("2");
        ws.GetRange("B7").SetValue("b");
        ws.GetRange("C7").SetValue("2");
        ws.GetRange("A8").SetValue("1");
        ws.GetRange("B8").SetValue("c");
        ws.GetRange("C8").SetValue("1");

        // Sort by A asc, then B asc, then C desc
        ws.GetRange("A1:C8").SetSort("A1", "xlAscending", "B1", "xlAscending", "C1", "xlDescending", "xlNo", "xlSortColumns");

        // Expected order:
        // (1,c,2), (1,c,1), (1,d,1),
        // (2,a,5), (2,a,1), (2,b,3), (2,b,2),
        // (3,e,9)

        // Row 1
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1", "R1 A");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "c", "R1 B");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "2", "R1 C");

        // Row 2
        assert.strictEqual(ws.GetRange("A2").GetValue(), "1", "R2 A");
        assert.strictEqual(ws.GetRange("B2").GetValue(), "c", "R2 B");
        assert.strictEqual(ws.GetRange("C2").GetValue(), "1", "R2 C");

        // Row 3
        assert.strictEqual(ws.GetRange("A3").GetValue(), "1", "R3 A");
        assert.strictEqual(ws.GetRange("B3").GetValue(), "d", "R3 B");
        assert.strictEqual(ws.GetRange("C3").GetValue(), "1", "R3 C");

        // Row 4
        assert.strictEqual(ws.GetRange("A4").GetValue(), "2", "R4 A");
        assert.strictEqual(ws.GetRange("B4").GetValue(), "a", "R4 B");
        assert.strictEqual(ws.GetRange("C4").GetValue(), "5", "R4 C");

        // Row 5
        assert.strictEqual(ws.GetRange("A5").GetValue(), "2", "R5 A");
        assert.strictEqual(ws.GetRange("B5").GetValue(), "a", "R5 B");
        assert.strictEqual(ws.GetRange("C5").GetValue(), "1", "R5 C");

        // Row 6
        assert.strictEqual(ws.GetRange("A6").GetValue(), "2", "R6 A");
        assert.strictEqual(ws.GetRange("B6").GetValue(), "b", "R6 B");
        assert.strictEqual(ws.GetRange("C6").GetValue(), "3", "R6 C");

        // Row 7
        assert.strictEqual(ws.GetRange("A7").GetValue(), "2", "R7 A");
        assert.strictEqual(ws.GetRange("B7").GetValue(), "b", "R7 B");
        assert.strictEqual(ws.GetRange("C7").GetValue(), "2", "R7 C");

        // Row 8
        assert.strictEqual(ws.GetRange("A8").GetValue(), "3", "R8 A");
        assert.strictEqual(ws.GetRange("B8").GetValue(), "e", "R8 B");
        assert.strictEqual(ws.GetRange("C8").GetValue(), "9", "R8 C");
    });

    QUnit.test("Multi-key sorting by A then C (A asc, C desc) with key2=null", function (assert) {
        // Initial data (A,B,C). B exists but must NOT be used as a tie-breaker.
        // r1: (2,b,3)
        // r2: (1,d,1)
        // r3: (2,a,5)
        // r4: (1,c,2)
        // r5: (2,z,2)
        // r6: (1,a,4)
        ws.GetRange("A1").SetValue("2");
        ws.GetRange("B1").SetValue("b");
        ws.GetRange("C1").SetValue("3");
        ws.GetRange("A2").SetValue("1");
        ws.GetRange("B2").SetValue("d");
        ws.GetRange("C2").SetValue("1");
        ws.GetRange("A3").SetValue("2");
        ws.GetRange("B3").SetValue("a");
        ws.GetRange("C3").SetValue("5");
        ws.GetRange("A4").SetValue("1");
        ws.GetRange("B4").SetValue("c");
        ws.GetRange("C4").SetValue("2");
        ws.GetRange("A5").SetValue("2");
        ws.GetRange("B5").SetValue("z");
        ws.GetRange("C5").SetValue("2");
        ws.GetRange("A6").SetValue("1");
        ws.GetRange("B6").SetValue("a");
        ws.GetRange("C6").SetValue("4");

        // Sort by A ascending, skip key2 (null), then by C descending.
        ws.GetRange("A1:C6").SetSort("A1", "xlAscending", null, null, "C1", "xlDescending", "xlNo", "xlSortColumns");

        // Expected order (A asc, then C desc; B ignored as a tie-breaker):
        // (1, a, 4), (1, c, 2), (1, d, 1), (2, a, 5), (2, b, 3), (2, z, 2)

        // Row 1
        assert.strictEqual(ws.GetRange("A1").GetValue(), "1", "R1 A");
        assert.strictEqual(ws.GetRange("B1").GetValue(), "a", "R1 B");
        assert.strictEqual(ws.GetRange("C1").GetValue(), "4", "R1 C");

        // Row 2
        assert.strictEqual(ws.GetRange("A2").GetValue(), "1", "R2 A");
        assert.strictEqual(ws.GetRange("B2").GetValue(), "c", "R2 B");
        assert.strictEqual(ws.GetRange("C2").GetValue(), "2", "R2 C");

        // Row 3
        assert.strictEqual(ws.GetRange("A3").GetValue(), "1", "R3 A");
        assert.strictEqual(ws.GetRange("B3").GetValue(), "d", "R3 B");
        assert.strictEqual(ws.GetRange("C3").GetValue(), "1", "R3 C");

        // Row 4
        assert.strictEqual(ws.GetRange("A4").GetValue(), "2", "R4 A");
        assert.strictEqual(ws.GetRange("B4").GetValue(), "a", "R4 B");
        assert.strictEqual(ws.GetRange("C4").GetValue(), "5", "R4 C");

        // Row 5
        assert.strictEqual(ws.GetRange("A5").GetValue(), "2", "R5 A");
        assert.strictEqual(ws.GetRange("B5").GetValue(), "b", "R5 B");
        assert.strictEqual(ws.GetRange("C5").GetValue(), "3", "R5 C");

        // Row 6
        assert.strictEqual(ws.GetRange("A6").GetValue(), "2", "R6 A");
        assert.strictEqual(ws.GetRange("B6").GetValue(), "z", "R6 B");
        assert.strictEqual(ws.GetRange("C6").GetValue(), "2", "R6 C");
    });

    QUnit.test("SetSort: A1:C4, sort by 'super' defined name, headers no, orientation column, should sort by A", function(assert) {
        // Fill A1:C4 with values:
        // 1 4 3
        // 2 3 4
        // 3 2 1
        // 4 1 2
        ws.GetRange("A1").SetValue("1"); ws.GetRange("B1").SetValue("4"); ws.GetRange("C1").SetValue("3");
        ws.GetRange("A2").SetValue("2"); ws.GetRange("B2").SetValue("3"); ws.GetRange("C2").SetValue("4");
        ws.GetRange("A3").SetValue("3"); ws.GetRange("B3").SetValue("2"); ws.GetRange("C3").SetValue("1");
        ws.GetRange("A4").SetValue("4"); ws.GetRange("B4").SetValue("1"); ws.GetRange("C4").SetValue("2");

        // Sort by key range C1:A1 (i.e., sort columns by row 1 descending)
        ws.GetRange("A1:C4").SetSort("super", "xlDescending", null, null, null, null, "xlNo", "xlSortColumns");

        // sort by A Column
        assert.strictEqual(ws.GetRange("A1").GetValue(), "4"); assert.strictEqual(ws.GetRange("B1").GetValue(), "1"); assert.strictEqual(ws.GetRange("C1").GetValue(), "2");
        assert.strictEqual(ws.GetRange("A2").GetValue(), "3"); assert.strictEqual(ws.GetRange("B2").GetValue(), "2"); assert.strictEqual(ws.GetRange("C2").GetValue(), "1");
        assert.strictEqual(ws.GetRange("A3").GetValue(), "2"); assert.strictEqual(ws.GetRange("B3").GetValue(), "3"); assert.strictEqual(ws.GetRange("C3").GetValue(), "4");
        assert.strictEqual(ws.GetRange("A4").GetValue(), "1"); assert.strictEqual(ws.GetRange("B4").GetValue(), "4"); assert.strictEqual(ws.GetRange("C4").GetValue(), "3");
    });

    QUnit.test("SetSort: negativeIndexColumn", function(assert) {
        ws.GetRange("F8").SetValue("1"); ws.GetRange('G8').SetValue("4"); ws.GetRange("H8").SetValue("3");
        ws.GetRange("F9").SetValue("2"); ws.GetRange('G9').SetValue("3"); ws.GetRange("H9").SetValue("4");
        ws.GetRange("F10").SetValue("3"); ws.GetRange('G10').SetValue("2"); ws.GetRange("H10").SetValue("1");
        ws.GetRange("F11").SetValue("4"); ws.GetRange('G11').SetValue("1"); ws.GetRange("H11").SetValue("2");

        ws.GetRange("G8:H11").SetSort("negativeIndexColumn", "xlDescending", null, null, null, null, "xlNo", "xlSortColumns");
        // should be no changes at all
        assert.strictEqual(ws.GetRange("F8").GetValue(), "1"); assert.strictEqual(ws.GetRange('G8').GetValue(), "4"); assert.strictEqual(ws.GetRange("H8").GetValue(), "3");
        assert.strictEqual(ws.GetRange("F9").GetValue(), "2"); assert.strictEqual(ws.GetRange('G9').GetValue(), "3"); assert.strictEqual(ws.GetRange("H9").GetValue(), "4");
        assert.strictEqual(ws.GetRange("F10").GetValue(), "3"); assert.strictEqual(ws.GetRange('G10').GetValue(), "2"); assert.strictEqual(ws.GetRange("H10").GetValue(), "1");
        assert.strictEqual(ws.GetRange("F11").GetValue(), "4"); assert.strictEqual(ws.GetRange('G11').GetValue(), "1"); assert.strictEqual(ws.GetRange("H11").GetValue(), "2");
    });

    QUnit.test("SetSort: negativeIndexRow", function(assert) {
        ws.GetRange("F8").SetValue("1"); ws.GetRange('G8').SetValue("4"); ws.GetRange("H8").SetValue("3");
        ws.GetRange("F9").SetValue("2"); ws.GetRange('G9').SetValue("3"); ws.GetRange("H9").SetValue("4");
        ws.GetRange("F10").SetValue("3"); ws.GetRange('G10').SetValue("2"); ws.GetRange("H10").SetValue("1");
        ws.GetRange("F11").SetValue("4"); ws.GetRange('G11').SetValue("1"); ws.GetRange("H11").SetValue("2");

        ws.GetRange("F9:H11").SetSort("negativeIndexRow", "xlDescending", null, null, null, null, "xlNo", "xlSortRows");
        // should be no changes at all
        assert.strictEqual(ws.GetRange("F8").GetValue(), "1"); assert.strictEqual(ws.GetRange('G8').GetValue(), "4"); assert.strictEqual(ws.GetRange("H8").GetValue(), "3");
        assert.strictEqual(ws.GetRange("F9").GetValue(), "2"); assert.strictEqual(ws.GetRange('G9').GetValue(), "3"); assert.strictEqual(ws.GetRange("H9").GetValue(), "4");
        assert.strictEqual(ws.GetRange("F10").GetValue(), "3"); assert.strictEqual(ws.GetRange('G10').GetValue(), "2"); assert.strictEqual(ws.GetRange("H10").GetValue(), "1");
        assert.strictEqual(ws.GetRange("F11").GetValue(), "4"); assert.strictEqual(ws.GetRange('G11').GetValue(), "1"); assert.strictEqual(ws.GetRange("H11").GetValue(), "2");
    });

    QUnit.test("SetSort: outOfRangeColumn", function(assert) {
        ws.GetRange("F8").SetValue("1"); ws.GetRange('G8').SetValue("4"); ws.GetRange("H8").SetValue("3");
        ws.GetRange("F9").SetValue("2"); ws.GetRange('G9').SetValue("3"); ws.GetRange("H9").SetValue("4");
        ws.GetRange("F10").SetValue("3"); ws.GetRange('G10').SetValue("2"); ws.GetRange("H10").SetValue("1");
        ws.GetRange("F11").SetValue("4"); ws.GetRange('G11').SetValue("1"); ws.GetRange("H11").SetValue("2");

        ws.GetRange("F8:G11").SetSort("outOfRangeColumn", "xlDescending", null, null, null, null, "xlNo", "xlSortColumns");
        // should be no changes at all
        assert.strictEqual(ws.GetRange("F8").GetValue(), "1"); assert.strictEqual(ws.GetRange('G8').GetValue(), "4"); assert.strictEqual(ws.GetRange("H8").GetValue(), "3");
        assert.strictEqual(ws.GetRange("F9").GetValue(), "2"); assert.strictEqual(ws.GetRange('G9').GetValue(), "3"); assert.strictEqual(ws.GetRange("H9").GetValue(), "4");
        assert.strictEqual(ws.GetRange("F10").GetValue(), "3"); assert.strictEqual(ws.GetRange('G10').GetValue(), "2"); assert.strictEqual(ws.GetRange("H10").GetValue(), "1");
        assert.strictEqual(ws.GetRange("F11").GetValue(), "4"); assert.strictEqual(ws.GetRange('G11').GetValue(), "1"); assert.strictEqual(ws.GetRange("H11").GetValue(), "2");
    });

    QUnit.test("SetSort: outOfRangeRow", function(assert) {
        ws.GetRange("F8").SetValue("1"); ws.GetRange('G8').SetValue("4"); ws.GetRange("H8").SetValue("3");
        ws.GetRange("F9").SetValue("2"); ws.GetRange('G9').SetValue("3"); ws.GetRange("H9").SetValue("4");
        ws.GetRange("F10").SetValue("3"); ws.GetRange('G10').SetValue("2"); ws.GetRange("H10").SetValue("1");
        ws.GetRange("F11").SetValue("4"); ws.GetRange('G11').SetValue("1"); ws.GetRange("H11").SetValue("2");

        ws.GetRange("F8:H10").SetSort("outOfRangeRow", "xlDescending", null, null, null, null, "xlNo", "xlSortRows");
        // should be no changes at all
        assert.strictEqual(ws.GetRange("F8").GetValue(), "1"); assert.strictEqual(ws.GetRange('G8').GetValue(), "4"); assert.strictEqual(ws.GetRange("H8").GetValue(), "3");
        assert.strictEqual(ws.GetRange("F9").GetValue(), "2"); assert.strictEqual(ws.GetRange('G9').GetValue(), "3"); assert.strictEqual(ws.GetRange("H9").GetValue(), "4");
        assert.strictEqual(ws.GetRange("F10").GetValue(), "3"); assert.strictEqual(ws.GetRange('G10').GetValue(), "2"); assert.strictEqual(ws.GetRange("H10").GetValue(), "1");
        assert.strictEqual(ws.GetRange("F11").GetValue(), "4"); assert.strictEqual(ws.GetRange('G11').GetValue(), "1"); assert.strictEqual(ws.GetRange("H11").GetValue(), "2");
    });
});
