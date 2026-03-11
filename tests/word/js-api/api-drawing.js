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

$(function ()
{
	QUnit.module("ApiDrawing");

    QUnit.test("GetFlipH", function (assert) {
        let docContent = AscTest.JsApi.CreateDocContent();
        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        let p = docContent.GetElement(0);
        p.AddDrawing(drawing);
        assert.strictEqual(drawing.GetFlipH(), false, 'Check drawing horizontal flip === false');
        drawing.SetHorFlip(true);
        assert.strictEqual(drawing.GetFlipH(), true, 'Check drawing horizontal flip === true');
    });

    QUnit.test("GetFlipV", function (assert) {
        let docContent = AscTest.JsApi.CreateDocContent();
        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        let p = docContent.GetElement(0);
        p.AddDrawing(drawing);
        assert.strictEqual(drawing.GetFlipV(), false, 'Check drawing vertical flip === false');
        drawing.SetVertFlip(true);
        assert.strictEqual(drawing.GetFlipV(), true, 'Check drawing vertical flip === true');
    });

    QUnit.test("CreateStroke", function (assert) {
        let docContent = AscTest.JsApi.CreateDocContent();
        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(3200*2, AscTest.JsApi.CreateNoFill(), "dash");
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        let p = docContent.GetElement(0);
        p.AddDrawing(drawing);
        assert.strictEqual(stroke.Ln.prstDash, 0, 'Check stroke dash type === 0');
    });

    QUnit.test("SetRelativeHeight", function (assert) {
        let docContent = AscTest.JsApi.CreateDocContent();
        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(3200*2, AscTest.JsApi.CreateNoFill(), "dash");
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        let p = docContent.GetElement(0);
        p.AddDrawing(drawing);
        drawing.SetRelativeHeight("page", 20);
        assert.equal(JSON.parse(drawing.ToJSON()).sizeRelV.relativeFrom, "page", 'Check drawing height relativeFrom === "page"');
        assert.equal(JSON.parse(drawing.ToJSON()).sizeRelV['wp14:pctHeight'], 20, 'Check drawing height pctHeight === 20');
    });

    QUnit.test("SetRelativeWidth", function (assert) {
        let docContent = AscTest.JsApi.CreateDocContent();
        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(3200*2, AscTest.JsApi.CreateNoFill(), "dash");
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        let p = docContent.GetElement(0);
        p.AddDrawing(drawing);
        drawing.SetRelativeWidth("page", 10);
        assert.equal(JSON.parse(drawing.ToJSON()).sizeRelH.relativeFrom, "page", 'Check drawing width relativeFrom === "page"');
        assert.equal(JSON.parse(drawing.ToJSON()).sizeRelH['wp14:pctWidth'], 10, 'Check drawing width pctWidth === 10');
    });

    QUnit.test("SetHorPosition", function (assert) {
        let docContent = AscTest.JsApi.CreateDocContent();
        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(3200*2, AscTest.JsApi.CreateNoFill(), "dash");
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        let p = docContent.GetElement(0);
        p.AddDrawing(drawing);
        drawing.SetHorPosition("page", 10, true);
        let oPositionH = JSON.parse(drawing.ToJSON()).positionH;
        assert.equal(oPositionH.relativeFrom, "page", 'Check drawing horizontal position relativeFrom === "page"');
        assert.equal(oPositionH.posOffset, 10, 'Check drawing horizontal position posOffset === 10');
        assert.equal(oPositionH.percent, true, 'Check drawing horizontal position percent === true');
    });

    QUnit.test("SetVerPosition", function (assert) {
        let docContent = AscTest.JsApi.CreateDocContent();
        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(3200*2, AscTest.JsApi.CreateNoFill(), "dash");
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        let p = docContent.GetElement(0);
        p.AddDrawing(drawing);
        drawing.SetVerPosition("topMargin", 20, true);
        let oPositionV = JSON.parse(drawing.ToJSON()).positionV;
        assert.equal(oPositionV.relativeFrom, "topMargin", 'Check drawing vertical position relativeFrom === "topMargin"');
        assert.equal(oPositionV.posOffset, 20, 'Check drawing vertical position posOffset === 20');
        assert.equal(oPositionV.percent, true, 'Check drawing vertical position percent === true');
    });

    QUnit.test("GetName", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);
        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);
        let name = drawing.GetName();
        assert.strictEqual(typeof name, 'string', 'Check GetName returns a string');
        assert.ok(name.length > 0, 'Check drawing name is not empty');
    });

    QUnit.test("SetName", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        let result = drawing.SetName("TestShape");
        assert.strictEqual(result, true, 'Check SetName returns true');
        assert.strictEqual(drawing.GetName(), "TestShape", 'Check drawing name is set correctly');

        let result2 = drawing.SetName("");
        assert.strictEqual(result2, false, 'Check SetName returns false for empty string');

        let result3 = drawing.SetName(null);
        assert.strictEqual(result3, false, 'Check SetName returns false for null');

        let result4 = drawing.SetName(undefined);
        assert.strictEqual(result4, false, 'Check SetName returns false for undefined');

        // Test that setting duplicate name causes previous drawing to get default name
        let drawing2 = AscTest.JsApi.CreateShape("rect", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing2);

        drawing.SetName("DuplicateName");
        const firstDrawingName = drawing.GetName();
        assert.strictEqual(firstDrawingName, "DuplicateName", 'Check first drawing has duplicate name');

        drawing2.SetName("DuplicateName");

        assert.strictEqual(drawing2.GetName(), "DuplicateName", 'Check second drawing has the duplicate name');
        assert.notStrictEqual(drawing.GetName(), "DuplicateName", 'Check first drawing name changed from duplicate');
        assert.notStrictEqual(drawing.GetName(), firstDrawingName, 'Check first drawing has a new default name');
    });

    QUnit.test("Unselect", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        drawing.Select();
        assert.ok(drawing.Drawing.getDrawingObjectsController().selectedObjects.includes(drawing.Drawing), 'Check drawing is selected before unselect');

        let result = drawing.Unselect();
        assert.strictEqual(result, true, 'Check Unselect returns true');
        assert.ok(!drawing.Drawing.getDrawingObjectsController().selectedObjects.includes(drawing.Drawing), 'Check drawing is not selected after unselect');
    });

    QUnit.test("Select", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        let result = drawing.Select();
        assert.strictEqual(result, true, 'Check Select returns true');
        assert.ok(drawing.Drawing.getDrawingObjectsController().selectedObjects.includes(drawing.Drawing), 'Check drawing is selected in document');
    });

    QUnit.test("SetTitle", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        let result = drawing.SetTitle("Test Title");
        assert.strictEqual(result, true, 'Check SetTitle returns true');
        assert.strictEqual(drawing.GetTitle(), "Test Title", 'Check title is set correctly');

        result = drawing.SetTitle("");
        assert.strictEqual(result, false, 'Check SetTitle returns false for empty string');

        result = drawing.SetTitle(null);
        assert.strictEqual(result, false, 'Check SetTitle returns false for null');

        result = drawing.SetTitle(undefined);
        assert.strictEqual(result, false, 'Check SetTitle returns false for undefined');
    });

    QUnit.test("GetTitle", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        assert.strictEqual(drawing.GetTitle(), null, 'Check GetTitle returns null when no title is set');

        drawing.SetTitle("My Title");
        assert.strictEqual(drawing.GetTitle(), "My Title", 'Check GetTitle returns the set title');
    });

    QUnit.test("SetDescription", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        let result = drawing.SetDescription("Test Description");
        assert.strictEqual(result, true, 'Check SetDescription returns true');
        assert.strictEqual(drawing.GetDescription(), "Test Description", 'Check description is set correctly');

        result = drawing.SetDescription("");
        assert.strictEqual(result, false, 'Check SetDescription returns false for empty string');

        result = drawing.SetDescription(null);
        assert.strictEqual(result, false, 'Check SetDescription returns false for null');

        result = drawing.SetDescription(undefined);
        assert.strictEqual(result, false, 'Check SetDescription returns false for undefined');
    });

    QUnit.test("GetDescription", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        assert.strictEqual(drawing.GetDescription(), null, 'Check GetDescription returns null when no description is set');

        drawing.SetDescription("My Description");
        assert.strictEqual(drawing.GetDescription(), "My Description", 'Check GetDescription returns the set description');
    });

    QUnit.test("SetLockAspect", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        assert.strictEqual(drawing.SetLockAspect(true), true, 'Check SetLockAspect returns true for valid value');
        assert.strictEqual(drawing.GetLockAspect(), true, 'Check lock aspect is set to true');

        assert.strictEqual(drawing.SetLockAspect(false), true, 'Check SetLockAspect returns true for false');
        assert.strictEqual(drawing.GetLockAspect(), false, 'Check lock aspect is set to false');

        assert.strictEqual(drawing.SetLockAspect(null), false, 'Check SetLockAspect returns false for null');
        assert.strictEqual(drawing.SetLockAspect(undefined), false, 'Check SetLockAspect returns false for undefined');
        assert.strictEqual(drawing.SetLockAspect(1), false, 'Check SetLockAspect returns false for non-boolean');
    });

    QUnit.test("GetLockAspect", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        assert.strictEqual(typeof drawing.GetLockAspect(), 'boolean', 'Check GetLockAspect returns a boolean');

        drawing.SetLockAspect(true);
        assert.strictEqual(drawing.GetLockAspect(), true, 'Check GetLockAspect returns true after SetLockAspect(true)');

        drawing.SetLockAspect(false);
        assert.strictEqual(drawing.GetLockAspect(), false, 'Check GetLockAspect returns false after SetLockAspect(false)');
    });

    QUnit.test("SetAllowOverlap", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        drawing.SetAllowOverlap(true);
        assert.strictEqual(drawing.GetAllowOverlap(), true, 'Check allow overlap is set to true');

        drawing.SetAllowOverlap(false);
        assert.strictEqual(drawing.GetAllowOverlap(), false, 'Check allow overlap is set to false');
    });

    QUnit.test("GetAllowOverlap", function (assert) {
        let doc = AscTest.JsApi.GetDocument();
        let p = AscTest.JsApi.CreateParagraph();
        doc.Push(p);

        let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.RGB(255, 111, 61));
        let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
        let drawing = AscTest.JsApi.CreateShape("cube", 3212465, 963295, fill, stroke);
        p.AddDrawing(drawing);

        assert.strictEqual(typeof drawing.GetAllowOverlap(), 'boolean', 'Check GetAllowOverlap returns a boolean');

        drawing.SetAllowOverlap(true);
        assert.strictEqual(drawing.GetAllowOverlap(), true, 'Check GetAllowOverlap returns true after SetAllowOverlap(true)');

        drawing.SetAllowOverlap(false);
        assert.strictEqual(drawing.GetAllowOverlap(), false, 'Check GetAllowOverlap returns false after SetAllowOverlap(false)');
    });
});

