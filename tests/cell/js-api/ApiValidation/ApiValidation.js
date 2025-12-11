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
    // ======= MUST-HAVE RUNTIME SETTINGS (kept exactly; do not remove) =======
    Asc.spreadsheet_api.prototype._init = function () {};
    Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
        callback();
    };
    AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {};
    AscCommonExcel.WorkbookView.prototype._init = function () {};
    AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {};
    AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {};
    AscCommonExcel.WorksheetView.prototype._init = function () {};
    AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable = function () {};
    AscCommonExcel.WorksheetView.prototype.setSelection = function () {};
    AscCommonExcel.WorksheetView.prototype.draw = function () {};
    AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {};
    AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () {};
    AscCommonExcel.WorksheetView.prototype.getZoom = function () {};
    AscCommonExcel.WorksheetView.prototype._getPPIY = function () {};
    AscCommonExcel.WorksheetView.prototype._getPPIX = function () {};
    AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {};
    Asc.ReadDefTableStyles = function () {};

    var api = new Asc.spreadsheet_api({
        "id-view": "editor_sdk",
    });

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
    wsView.objectRender.controller = new AscFormat.DrawingObjectsController(wsView.objectRender);
    var ws = api.GetActiveSheet();

    // ======= TEST HELPERS =======
    // Must exist & each test must start with it
    const initializeTest = function (/*rangeAddress optional*/) {
        const globalRange = ws.GetRange('A1:Z100'); // acceptable sandbox
        globalRange.Clear();
        // Reset validations entirely
        if (ws.worksheet && ws.worksheet.dataValidations) {
            ws.worksheet.dataValidations.clear(ws.worksheet, true);
        }
    };
    window.initializeTest = initializeTest; // expose for debugging if needed

    const isSameRange = function (rA, rB) {
        return rA && rB && (rA.r1 === rB.r1 && rA.c1 === rB.c1 && rA.r2 === rB.r2 && rA.c2 === rB.c2);
    };

    const getAllValidations = function () {
        return (ws.worksheet && ws.worksheet.dataValidations && ws.worksheet.dataValidations.elems) || [];
    };

    const countValidations = function () {
        return getAllValidations().length;
    };

    const a1ToRC0 = function (addr) {
        // crude: letters -> col, digits -> row (1-based)
        const target = ws.GetRange(addr); // to ensure address is valid
        if (!target) return null;

        return { r: target.range.bbox.r1, c: target.range.bbox.c1 };
    };

    const cellInAnyValidation = function (a1) {
        const target = a1ToRC0(a1);
        if (!target) return false;
        const { r, c } = target;
        const all = getAllValidations();
        for (let i = 0; i < all.length; i++) {
            const ranges = all[i].ranges || [];
            for (let j = 0; j < ranges.length; j++) {
                if (ranges[j].contains(c, r)) return true;
            }
        }
        return false;
    };

    const findRangeRecord = function (bbox) {
        const all = getAllValidations();
        for (let i = 0; i < all.length; i++) {
            const ranges = all[i].ranges || [];
            for (let j = 0; j < ranges.length; j++) {
                if (isSameRange(ranges[j], bbox)) return { v: all[i], range: ranges[j] };
            }
        }
        return null;
    };

    // ======= TESTS =======

    QUnit.module("ApiValidation — smoke & helpers");

    QUnit.test("Empty validation placeholder on a fresh cell", function (assert) {
        initializeTest();
        const v = ws.GetRange("A1").GetValidation();
        assert.ok(v, "Validation object exists");
        assert.ok(Array.isArray(v.validations), "validations array exists");
        assert.strictEqual(v.validations.length, 1, "One placeholder validation exists");
        assert.strictEqual(v.validations[0].ranges, null, "Placeholder validation has null ranges");
    });

    QUnit.module("ApiValidation.Add — basics, formulas, multi-area, overlap");

    QUnit.test("Add simple numeric between on A1:C3", function (assert) {
        initializeTest();
        const r = ws.GetRange("A1:C3");
        r.SetValue("10");
        const res = r.GetValidation().Add('xlValidateDecimal', 'xlValidAlertWarning', 'xlBetween', '5', '15');
        assert.ok(res, "Add returns instance");
        assert.strictEqual(countValidations(), 1, "One validation stored");
        assert.ok(isSameRange(getAllValidations()[0].ranges[0], r.range.bbox), "Range bbox matches");
    });

    QUnit.test("Add with formulas: string, number, ApiRange", function (assert) {
        initializeTest();

        // string
        const r1 = ws.GetRange("D1:D2");
        r1.GetValidation().Add('xlValidateList', 'xlValidAlertStop', undefined, 'A1:A3');
        assert.strictEqual(r1.GetValidation().GetFormula1(), '\"A1:A3\"', "Formula from string is quoted");

        // number
        const r2 = ws.GetRange("E1:E1");
        r2.GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 5, 10);
        assert.strictEqual(r2.GetValidation().GetFormula1(), '5', "Formula1 from number");
        assert.strictEqual(r2.GetValidation().GetFormula2(), '10', "Formula2 from number");

        // ApiRange
        const src = ws.GetRange("F1:F2"); src.SetValue('x');
        const r3 = ws.GetRange("G1:G2");
        r3.GetValidation().Add('xlValidateList', 'xlValidAlertWarning', undefined, ws.GetRange('F1:F2'));
        assert.strictEqual(r3.GetValidation().GetFormula1(), '\"F1:F2\"', "Formula1 from ApiRange address");
    });

    QUnit.test("Add fails when type is invalid", function (assert) {
        initializeTest();
        const r = ws.GetRange("A5:A6");
        const before = countValidations();
        const res = r.GetValidation().Add('xlValidate__Bogus', 'xlValidAlertStop', 'xlBetween', 1, 2);
        assert.strictEqual(res, null, "Add returns null for invalid type");
        assert.strictEqual(countValidations(), before, "No new validations added");
    });

    QUnit.test("Add should not work on an area overlapping an existing validation (contains & intersects)", function (assert) {
        initializeTest();
        // Seed a validation on D1:D5
        ws.GetRange("D1:D5").GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 3);
        const before = countValidations();

        // 1) 'Contains': try to add inside existing (D2:D4)
        const inside = ws.GetRange("D2:D4");
        const resInside = inside.GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 5, 6);
        const afterInside = countValidations();
        
        // 2) 'Intersects': try to add partially crossing (C4:E7 intersects D1:D5 on D4:D5)
        const inter = ws.GetRange("C4:E7");
        const resInter = inter.GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 7, 8);
        const afterInter = countValidations();

        // Implementation detail: Add() must return null and keep validation count unchanged
        assert.ok(resInside === null && afterInside === before, "Add inside overlapping area is blocked");
        assert.ok(resInter === null && afterInter === afterInside, "Add intersecting overlapping area is blocked");
    });

    QUnit.test("Add allowed on disjoint area creates a second validation", function (assert) {
        initializeTest();
        ws.GetRange("A1:A3").GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 2);
        const before = countValidations();
        ws.GetRange("C1:C3").GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 2);
        assert.strictEqual(countValidations(), before + 1, "Second disjoint validation created");
    });

    QUnit.module("ApiValidation getters/setters — coverage");

    QUnit.test("Type, Operator, AlertStyle, flags and titles/messages, Parent, Value mapping", function (assert) {
        initializeTest();
        const r = ws.GetRange("H1:H3");
        r.GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 100);

        const v = r.GetValidation();
        // Type
        assert.strictEqual(v.GetType(), 'xlValidateWholeNumber', "GetType matches");
        v.SetType('xlValidateDecimal');
        assert.strictEqual(v.GetType(), 'xlValidateDecimal', "SetType updated");

        // Operator
        assert.strictEqual(v.GetOperator(), 'xlBetween', "Default operator between");
        v.SetOperator('xlGreater');
        assert.strictEqual(v.GetOperator(), 'xlGreater', "Operator updated");

        // AlertStyle
        assert.strictEqual(v.GetAlertStyle(), 'xlValidAlertStop', "Default alert stop");
        v.SetAlertStyle('xlValidAlertWarning');
        assert.strictEqual(v.GetAlertStyle(), 'xlValidAlertWarning', "AlertStyle updated");

        // Flags
        v.SetIgnoreBlank(false);
        assert.strictEqual(v.GetIgnoreBlank(), false, "IgnoreBlank toggled");

        v.SetShowInput(false);
        assert.strictEqual(v.GetShowInput(), false, "ShowInput toggled");

        v.SetShowError(false);
        assert.strictEqual(v.GetShowError(), false, "ShowError toggled");

        // InCellDropdown inversion
        v.SetInCellDropdown(true);
        assert.strictEqual(v.GetInCellDropdown(), true, "InCellDropdown true after set");
        v.SetInCellDropdown(false);
        assert.strictEqual(v.GetInCellDropdown(), false, "InCellDropdown false after set");

        // Titles & messages
        v.SetInputTitle("Enter age");
        v.SetInputMessage("1..100 only");
        v.SetErrorTitle("Bad age");
        v.SetErrorMessage("Must be 1..100");

        assert.strictEqual(v.GetInputTitle(), "Enter age", "InputTitle OK");
        assert.strictEqual(v.GetInputMessage(), "1..100 only", "InputMessage OK");
        assert.strictEqual(v.GetErrorTitle(), "Bad age", "ErrorTitle OK");
        assert.strictEqual(v.GetErrorMessage(), "Must be 1..100", "ErrorMessage OK");

        // Formulas & Value property mapping
        v.SetFormula1("2");
        v.SetFormula2("99");
        assert.strictEqual(v.GetFormula1(), "2", "Formula1 OK");
        assert.strictEqual(v.GetFormula2(), "99", "Formula2 OK");
        v.Value = "10";
        assert.strictEqual(v.Value, "10", "Value <-> Formula1 mapping works");

        // Parent
        assert.ok(v.GetParent(), "Parent exists");
        assert.ok(v.Parent, "Parent property getter exists");
    });

    QUnit.module("ApiValidation.Delete — inside, contains, partial overlap");

    QUnit.test("Delete: range inside validation (remove the sub-area only)", function (assert) {
        initializeTest();
        // single validation on A5:C7
        const big = ws.GetRange("A5:C7");
        big.GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 10);
        assert.strictEqual(countValidations(), 1, "One validation to start");

        // delete B6 cell
        ws.GetRange("B6").GetValidation().Delete();

        assert.strictEqual(countValidations(), 1, "Validation object remains");
        assert.strictEqual(cellInAnyValidation("B6"), false, "B6 no longer validated");
        assert.strictEqual(cellInAnyValidation("A5"), true, "Non-deleted part still validated");
        assert.strictEqual(cellInAnyValidation("C7"), true, "Edge still validated");
    });

    QUnit.test("Delete: validation inside range (entire validation removed)", function (assert) {
        initializeTest();
        // validation at C10:D12
        ws.GetRange("C10:D12").GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 10);
        const before = countValidations();
        assert.ok(before >= 1, "At least one validation exists");

        // delete a superset A1:Z100
        ws.GetRange("A1:Z100").GetValidation().Delete();
        assert.strictEqual(countValidations(), 0, "All validations removed within big range");
        assert.strictEqual(cellInAnyValidation("C11"), false, "C11 not validated anymore");
    });

    QUnit.test("Delete: partial overlap (trim intersecting parts, keep the rest)", function (assert) {
        initializeTest();
        // validate A1:C3
        ws.GetRange("A1:C3").GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 9);
        assert.strictEqual(countValidations(), 1, "Seeded one validation");

        // delete B2:D4 (overlaps with B2:C3 area)
        ws.GetRange("B2:D4").GetValidation().Delete();

        // A1, A2 remain validated; B2,C3 become unvalidated
        assert.strictEqual(cellInAnyValidation("A1"), true, "A1 still validated");
        assert.strictEqual(cellInAnyValidation("A3"), true, "A3 still validated");
        assert.strictEqual(cellInAnyValidation("B2"), false, "B2 trimmed out");
        assert.strictEqual(cellInAnyValidation("C3"), false, "C3 trimmed out");
    });

    QUnit.module("ApiValidation.Modify — inside, contains/merge, partial overlap");

    QUnit.test("Modify: inside case (like delete sub-area then add new validation)", function (assert) {
        initializeTest();
        // Seed A5:C7
        ws.GetRange("A5:C7").GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 10);

        // Modify on B6:B7 — should carve out & create new rule there
        const m = ws.GetRange("B6:B7").GetValidation();
        m.Modify('xlValidateDecimal', 'xlValidAlertWarning', 'xlGreater', '5');

        // New validation should exist for B6:B7 with decimal/greater
        const rec = findRangeRecord(ws.GetRange("B6:B7").range.bbox);
        assert.ok(rec, "New validation covering B6:B7 exists");
        // Check type/operator by reading via a cell in that area
        const mv = ws.GetRange("B6:B7").GetValidation();
        assert.strictEqual(mv.GetType(), 'xlValidateDecimal', "Type changed for modified area");
        assert.strictEqual(mv.GetOperator(), 'xlGreater', "Operator changed for modified area");

        // Outside part (e.g., A5 still) remains validated by original
        assert.strictEqual(cellInAnyValidation("A5"), true, "Original area remains around modified carve");
    });

    QUnit.test("Modify: merge multiple existing validations into a single one on a bigger area", function (assert) {
        initializeTest();
        ws.GetRange('A10:A11').GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 5);
        ws.GetRange('C10:C11').GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 5);

        const mergeRange = ws.GetRange('A10:C11').GetValidation();
        assert.ok(mergeRange.validations && mergeRange.validations.length >= 2, "Selection sees multiple validations");
        mergeRange.Modify('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 2, 8);

        // Should be one validation that includes whole A10:C11
        const all = getAllValidations();
        assert.ok(all.length === 1, "Exactly one validation exists");
        const bbox = ws.GetRange('A10:C11').range.bbox;
        const merged = findRangeRecord(bbox);
        assert.ok(merged, "Merged validation covers A10:C11");
    });

    QUnit.test("Modify: partial overlap — trims originals and adds new one on the modified area", function (assert) {
        initializeTest();
        // Original on A1:C3
        ws.GetRange("A1:C3").GetValidation().Add('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 9);

        // Modify B2:D4
        ws.GetRange("B2:D4").GetValidation().Modify('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 3, 7);

        // There must be a validation range exactly for B2:D4 now
        const bbox = ws.GetRange("B2:D4").range.bbox;
        const rec = findRangeRecord(bbox);
        assert.ok(rec, "New modified validation exists for B2:D4");

        // Original trimmed part checks
        assert.strictEqual(cellInAnyValidation("A1"), true, "A1 kept by trimmed original");
        assert.strictEqual(cellInAnyValidation("B2"), true, "B2 now validated by new rule");
        assert.strictEqual(cellInAnyValidation("C3"), true, "C3 falls under new or trimmed original logic");
    });

    QUnit.module("ApiValidation edge-cases — safety");

    QUnit.test("Delete on area without validations: no throw, no changes", function (assert) {
        initializeTest();
        const before = countValidations();
        ws.GetRange("Z1:Z2").GetValidation().Delete();
        assert.strictEqual(countValidations(), before, "Nothing changed");
    });

    QUnit.test("Modify when there is no validation: returns null", function (assert) {
        initializeTest();
        const v = ws.GetRange("Y1:Y2").GetValidation();
        const res = v.Modify('xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', 1, 2);
        assert.strictEqual(res, null, "Modify returns null when nothing to modify");
    });

    QUnit.module("Enum conversion helpers — sanity");
});
