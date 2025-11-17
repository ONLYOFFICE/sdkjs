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

const DecToHex = function (n) {
    let hex = n.toString(16);
    return hex.length === 1 ? "0" + hex : hex;
};

$(function () {
    // ====== REQUIRED ENVIRONMENT SETUP (preserve these stubs/settings) ======
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

    var api = new Asc.spreadsheet_api({ "id-view": "editor_sdk" });

    api.FontLoader = { LoadDocumentFonts: function () {} };
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
        r.ClearFormats();
    };

    function color(r, g, b) {
        return api.CreateColorFromRGB(r, g, b);
    }
    function fullRange() {
        return ws.GetRange("A1:Z100");
    }

    // Seed some values that a few rules might reference
    (function seedValues() {
        var vals = [10, 20, 30, 40, 50];
        for (var i = 0; i < vals.length; i++) {
            ws.GetRange("A" + (i + 1)).SetValue(vals[i]);
            ws.GetRange("B" + (i + 1)).SetValue(vals[vals.length - 1 - i]);
        }
        ws.GetRange("C1").SetValue("Hello world");
        ws.GetRange("C2").SetValue("world");
        ws.GetRange("D1").SetValue(new Date()); // today
        // extra seeds
        for (var r = 1; r <= 10; r++) {
            ws.GetRange("E" + r).SetValue(r * 3);
            ws.GetRange("F" + r).SetValue(r % 2 === 0 ? "dup" : "uniq" + r);
            ws.GetRange("G" + r).SetValue(r);
            ws.GetRange("H" + r).SetValue(100 - r);
            ws.GetRange("I" + r).SetValue(r * -1);
        }
    })();

    // ====== TESTS ======

    QUnit.module("ApiFormatConditions", function () {
        QUnit.test("initializeTest clears CF in A1:Z100", function (assert) {
            initializeTest();
            var range = ws.GetRange("A1:A5");
            var fc = range.GetFormatConditions();
            var c = fc.Add("xlCellValue", "xlLess", "25");
            assert.ok(c, "Added initial CF");
            assert.strictEqual(fc.GetCount(), 1, "Exactly one CF added");
            // Now clear
            initializeTest();
            assert.strictEqual(
                fullRange().GetFormatConditions().GetCount(),
                0,
                "All CFs cleared by initializeTest"
            );
        });

        QUnit.test(
            "Add + GetItem + basic getters (xlCellValue)",
            function (assert) {
                initializeTest();
                var range = ws.GetRange("A1:A5");
                var fc = range.GetFormatConditions();
                var c1 = fc.Add("xlCellValue", "xlLess", "25");
                c1.SetFillColor(color(255, 0, 0));
                c1.SetNumberFormat("0.00");

                assert.ok(c1, "Condition created");
                assert.strictEqual(fc.GetCount(), 1, "GetCount returns 1");
                assert.ok(fc.GetItem(1), "GetItem(1) returns a condition");
                assert.strictEqual(c1.GetType(), "xlCellValue", "Type mapped");
                assert.strictEqual(c1.GetOperator(), "xlLess", "Operator mapped");
                assert.strictEqual(c1.GetFormula1(), "25", "Formula1 is set");
                assert.notStrictEqual(c1.GetFillColor(), "No Fill", "FillColor is applied");
                assert.strictEqual(c1.GetNumberFormat(), "0.00", "NumberFormat roundtrips");
                assert.ok(c1.GetFont(), "Font accessor returns a font object");
                assert.ok(c1.GetAppliesTo(), "AppliesTo is available");
                assert.strictEqual(c1.GetScopeType(), "xlSelectionScope", "Default scope is selection");
            }
        );

        QUnit.test(
            "Modify operator and formula; SetBorders; ScopeType switch",
            function (assert) {
                initializeTest();
                var range = ws.GetRange("A1:A5");
                var fc = range.GetFormatConditions();
                var c = fc.Add("xlCellValue", "xlBetween", "15", "35");
                assert.strictEqual(c.GetOperator(), "xlBetween", "Initial operator set");

                // Modify to greater than 15
                c.Modify(null, "xlGreater", "15");
                assert.strictEqual(c.GetOperator(), "xlGreater", "Operator changed via Modify");
                assert.strictEqual(c.GetFormula1(), "15", "Formula1 changed via Modify");

                // Modify using range formula (ensure address is used)
                var a1 = ws.GetRange("A1");
                c.Modify(null, "xlGreater", a1);
                assert.strictEqual(c.GetFormula1(), "A1", "Formula1 updated to range address");

                // Borders should not throw
                c.SetBorders("Top", "Continuous", color(0, 0, 0));
                assert.ok(true, "SetBorders executed");

                // ScopeType change
                c.SetScopeType("xlDataFieldScope");
                assert.strictEqual(c.GetScopeType(), "xlDataFieldScope", "ScopeType changed");
            }
        );

        QUnit.test(
            "Priority controls: SetFirstPriority / SetLastPriority / SetPriority",
            function (assert) {
                initializeTest();
                var range = ws.GetRange("A1:A5");
                var fc = range.GetFormatConditions();
                var c1 = fc.Add("xlCellValue", "xlGreater", "5");
                var c2 = fc.Add("xlCellValue", "xlLess", "50");
                var c3 = fc.Add("xlCellValue", "xlEqual", "30");

                // Priorities should be increasing
                var p1 = c1.GetPriority(),
                    p2 = c2.GetPriority(),
                    p3 = c3.GetPriority();
                assert.ok(p1 < p2 && p2 < p3, "Priorities assigned in increasing order" );

                // Make c2 first
                c2.SetFirstPriority();
                assert.strictEqual(c2.GetPriority(), 1, "c2 is now first priority");

                // Make c1 last
                c1.SetLastPriority();
                assert.ok(c1.GetPriority() > c3.GetPriority(), "c1 moved to the last priority");

                // Set an explicit unique priority
                var maxP = Math.max(
                    c1.GetPriority(),
                    c2.GetPriority(),
                    c3.GetPriority()
                );

                c3.SetPriority(maxP + 10);
                assert.strictEqual(
                    c3.GetPriority(),
                    maxP + 10,
                    "c3 explicit priority applied"
                );
            }
        );

        QUnit.test("Text rules: Text, TextOperator and generated formulas", function (assert) {
            initializeTest();
            var r = ws.GetRange("C1:C5");
            var fc = r.GetFormatConditions();
            var c = fc.Add("xlTextString", "xlBeginsWith", "Hel");
            assert.ok(c, "Text rule created");
            assert.strictEqual(c.GetTextOperator(), "xlBeginsWith", "Initial TextOperator mapped");


            c.SetText("world");
            assert.strictEqual(c.GetText(), "world", "Text property updated");
            assert.strictEqual(c.GetFormula1(), 'LEFT(C1,LEN(\"world\"))=\"world\"', "Formula1 generated correctly for BeginsWith");

            c.SetTextOperator("xlContains");
            assert.strictEqual(c.GetTextOperator(), "xlContains", "TextOperator changed");

            assert.strictEqual(c.GetFormula1(), "NOT(ISERROR(SEARCH(\"world\",C1)))", "Formula1 generated correctly for Contains");
        });

        QUnit.test("Time period rule: DateOperator get/set", function (assert) {
            initializeTest();
            var r = ws.GetRange("D1:D5");
            var fc = r.GetFormatConditions();
            var c = fc.Add("xlTimePeriod", "xlYesterday");
            assert.strictEqual(c.GetDateOperator(), "xlYesterday", "DateOperator mapped");

            c.SetDateOperator("xlToday");
            assert.strictEqual(c.GetDateOperator(), "xlToday", "DateOperator updated");
        });

        QUnit.test("ColorScale (2-color) criteria manipulation", function (assert) {
            initializeTest();
            var r = ws.GetRange("A1:A10");
            var fc = r.GetFormatConditions();
            var cs = fc.AddColorScale(2);
            assert.ok(cs, "2-color scale added");
            assert.strictEqual(cs.GetType(), "xlColorScale", "Type is xlColorScale");


            var criteria = cs.ColorScaleCriteria;
            assert.ok( criteria && criteria.length === 2, "Two criteria returned");

            criteria[0].SetType("xlConditionValueAutomaticMin");
            criteria[1].SetType("xlConditionValueAutomaticMax");
            criteria[0].SetValue("0"); // acceptable even if ignored for AutoMin internally
            criteria[1].SetValue("100");

            criteria[0].SetColor(color(248, 105, 107));
            criteria[1].SetColor(color(99, 190, 123));

            assert.strictEqual(
                criteria[0].GetType(),
                "xlConditionValueAutomaticMin",
                "First criterion type set"
            );
            assert.strictEqual(
                criteria[1].GetType(),
                "xlConditionValueAutomaticMax",
                "Second criterion type set"
            );
            // Indexes are ordered
            assert.ok(criteria[0].GetIndex() < criteria[1].GetIndex(), "Criterion indexes in order");
            assert.ok(criteria[0].GetColor(), "GetColor returns ApiColor");
        });

        QUnit.test( "ColorScale (3-color) criteria manipulation", function (assert) {
            initializeTest();
            var r = ws.GetRange("B1:B10");
            var fc = r.GetFormatConditions();
            var cs = fc.AddColorScale(3);
            var cr = cs.ColorScaleCriteria;
            assert.ok(cr && cr.length === 3, "Three criteria returned");

            cr[1].SetType("xlConditionValuePercentile");
            cr[1].SetValue("60");
            cr[1].SetColor(color(255, 235, 132));

            assert.strictEqual(cr[1].GetType(), "xlConditionValuePercentile", "Mid criterion percentile set");
            assert.strictEqual(cr[1].GetValue(), "60", "Mid criterion value set");
            assert.ok(cr[1].GetIndex() > cr[0].GetIndex() && cr[2].GetIndex() > cr[1].GetIndex(), "Indices 0 < 1 < 2");
        });

        QUnit.test("DataBar: axis, direction, show values, colors, min/max and lengths", function (assert) {
            initializeTest();
            var r = ws.GetRange("C1:C10");
            var fc = r.GetFormatConditions();
            var db = fc.AddDatabar();
            assert.ok(db, "Data bar created");
            assert.strictEqual(db.GetType(), "xlDatabar", "Type is xlDatabar");

            db.SetAxisPosition("xlDataBarAxisMidpoint");
            assert.strictEqual(
                db.GetAxisPosition(),
                "xlDataBarAxisMidpoint",
                "Axis position set/get"
            );

            db.SetDirection("xlRTL");
            assert.strictEqual(db.GetDirection(), "xlRTL", "Direction set/get");

            db.SetShowValue(false);
            assert.strictEqual(db.GetShowValue(), false, "ShowValue set/get");

            db.SetBarColor(color(99, 142, 198));
            db.SetBarBorderColor(color(10, 20, 30));
            db.NegativeBarColor = color(200, 50, 50);
            db.NegativeBorderColor = color(50, 50, 50);
            assert.ok(true, "Colors applied without error");

            db.PercentMin = 5;
            db.PercentMax = 95;
            assert.strictEqual(db.PercentMin, 5, "PercentMin set/get");
            assert.strictEqual(db.PercentMax, 95, "PercentMax set/get");

            // CFVO types and values
            db.SetMinPointType("xlConditionValueNumber");
            db.SetMinPointValue(10);
            db.SetMaxPointType("xlConditionValueNumber");
            db.SetMaxPointValue(90);
            assert.strictEqual(db.GetMinPointType(), "xlConditionValueNumber", "Min type set");
            assert.strictEqual(db.GetMaxPointType(), "xlConditionValueNumber", "Max type set");

            assert.strictEqual(db.GetMinPointValue(), "10", "Min value set");
            assert.strictEqual(db.GetMaxPointValue(), "90", "Max value set");
        });

        QUnit.test("IconSet: change set, percentile thresholds and reverse order", function (assert) {
            initializeTest();
            var r = ws.GetRange("D1:D10");
            var fc = r.GetFormatConditions();
            var ic = fc.AddIconSetCondition();
            assert.ok(ic, "Icon set condition created");

            // ShowIconOnly toggle
            ic.SetShowIconOnly(true);
            assert.strictEqual(ic.GetShowIconOnly(), true, "ShowIconOnly toggled");

            ic.SetIconSet("xl5Quarters");
            assert.strictEqual(ic.GetIconSet(), "xl5Quarters", "Icon set switched to 5 quarters");


            ic.SetPercentileValues(true);
            assert.strictEqual(ic.GetPercentileValues(), true, "PercentileValues enabled");

            ic.SetReverseOrder(true);
            assert.strictEqual(ic.GetReverseOrder(), true, "Reverse order enabled");

            // With 5 icons the criteria length should be 5
            var crit5 = ic.GetIconCriteria();
            assert.ok(crit5 && crit5.length === 5, "IconCriteria length matches icon set");

            // Switch to 3 icons and validate length
            ic.SetIconSet("xl3TrafficLights1");
            assert.strictEqual(ic.GetIconSet(), "xl3TrafficLights1", "Icon set switched to 3 traffic lights");

            // Change one criterion operator/value
            var crit3 = ic.GetIconCriteria();
            crit3[1].SetOperator("xlGreaterEqual");
            crit3[1].SetValue(50);
            assert.ok(crit3[1].GetValue() !== null, "Criterion value updated");
        });

        QUnit.test("AboveAverage: defaults and setters", function (assert) {
            initializeTest();
            var r = ws.GetRange("E1:E10");
            var fc = r.GetFormatConditions();
            var aa = fc.AddAboveAverage();
            assert.ok(aa, "AboveAverage rule created");
            assert.strictEqual(aa.GetType(), "xlAboveAverageCondition", "Type recognized");

            assert.strictEqual(aa.GetAboveBelow(), true, "Default AboveBelow is Above");

            aa.SetAboveBelow(false);
            assert.strictEqual(aa.GetAboveBelow(), false, "AboveBelow toggled to false");

            aa.SetNumStdDev(2);
            assert.strictEqual(aa.GetNumStdDev(), 2, "NumStdDev set to 2");
        });

        QUnit.test("Top10: defaults and setters (top/bottom, percent, rank)", function (assert) {
            initializeTest();
            var r = ws.GetRange("G1:G10");
            var fc = r.GetFormatConditions();
            var t = fc.AddTop10();
            assert.ok(t, "Top10 rule created");
            // defaults
            assert.strictEqual(t.GetType(), "xlTop10", "Type is xlTop10");
            assert.strictEqual(t.GetRank(), 10, "Default Rank is 10");
            assert.strictEqual(t.GetPercent(), false, "Default Percent is false");
            assert.strictEqual(t.GetTopBottom(), "xlTop10Top", "Default TopBottom is Top");

            // setters
            t.SetRank(3);
            t.SetPercent(true);
            t.SetTopBottom("xlTop10Bottom");
            assert.strictEqual(t.GetRank(), 3, "Rank updated");
            assert.strictEqual(t.GetPercent(), true, "Percent updated");
            assert.strictEqual(t.GetTopBottom(), "xlTop10Bottom", "TopBottom updated");

            // style APIs available
            t.SetFillColor(color(220, 230, 241));
            t.SetNumberFormat("0");
            assert.strictEqual(t.GetNumberFormat(), "0", "Top10 NumberFormat roundtrips");

            // applies-to expansion
            var rMulti = ws.GetRange("G1:G5,H1:H5");
            t.ModifyAppliesToRange(rMulti);
            assert.ok(t.GetAppliesTo(), "AppliesTo after modify");
        });

        QUnit.test("UniqueValues and Delete()", function (assert) {
            initializeTest();
            var r = ws.GetRange("F1:F10");
            var fc = r.GetFormatConditions();
            var u = fc.AddUniqueValues();
            assert.ok(u, "UniqueValues rule created");
            assert.strictEqual(fc.GetCount(), 1, "One rule present");

            // Add second rule then delete all
            fc.Add("xlCellValue", "xlGreater", "0");
            assert.strictEqual(fc.GetCount(), 2, "Two rules present");
            fc.Delete();
            assert.strictEqual(fc.GetCount(), 0, "Delete cleared all rules for the range");
        });

        QUnit.test("GetCount considers intersection only", function (assert) {
            initializeTest();
            // Add a rule outside A1:A5
            ws.GetRange("X1:X5").GetFormatConditions().Add("xlCellValue", "xlGreater", "0");
            // Query another separate range
            var c = ws.GetRange("A1:A5").GetFormatConditions().GetCount();
            assert.strictEqual(c, 0, "No intersecting rules -> count is 0");
        });

        QUnit.test("FormatCondition.Delete removes single rule", function (assert) {
            initializeTest();
            var r = ws.GetRange("I1:I10");
            var fc = r.GetFormatConditions();
            var c1 = fc.Add("xlCellValue", "xlGreater", "0");
            var c2 = fc.Add("xlCellValue", "xlLess", "0");
            assert.strictEqual(fc.GetCount(), 2, "Two rules present");
            c1.Delete();
            assert.strictEqual(fc.GetCount(), 1, "Single rule deleted");
        });

        QUnit.test("ApiDatabar extra props: AxisColor and BarFillType", function (assert) {
            initializeTest();
            var r = ws.GetRange("H1:H10");
            var fc = r.GetFormatConditions();
            var db = fc.AddDatabar();
            // AxisColor
            db.SetAxisColor(color(0, 0, 0));
            var axisColor = db.GetAxisColor();
            assert.ok(axisColor, "AxisColor set and retrieved");
            // BarFillType
            db.SetBarFillType("xlDataBarFillSolid");
            assert.strictEqual(db.GetBarFillType(), "xlDataBarFillSolid", "BarFillType set/get works");
        });

        QUnit.test("ApiFormatCondition with expression", function (assert) {
            initializeTest();
            var r = ws.GetRange("A1:A10");
            var fc = r.GetFormatConditions();
            var c = fc.Add("xlExpression", null, "=$A1>10");
            assert.strictEqual(c.Type, "xlExpression", "Expression rule created");
            assert.strictEqual(c.GetFormula1(), "$A1>10", "Expression formula kept");
        });
    });
});
