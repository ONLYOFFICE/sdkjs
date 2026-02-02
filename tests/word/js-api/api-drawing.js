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

});
