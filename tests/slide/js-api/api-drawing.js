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

"use strict";

$(function () {

    const logicDocument = AscTest.CreateLogicDocument();
	QUnit.module("ApiDrawing");

    function CreateSlide()
	{
		logicDocument.addNextSlide(0);
		editor.WordControl.Thumbnails.CalculatePlaces();
	}

	QUnit.test("Test: Create shape with gradient fill", function (assert) {
        CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);

		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		assert.ok(drawing, "Drawing should be created");

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		const gs1 = AscTest.JsApi.CreateGradientStop(AscTest.JsApi.CreateRGBColor(255, 213, 191), 0);
		const gs2 = AscTest.JsApi.CreateGradientStop(AscTest.JsApi.CreateRGBColor(255, 111, 61), 100000);
		const gradientFill = AscTest.JsApi.CreateRadialGradientFill([gs1, gs2]);

		drawing.Fill(gradientFill);

        assert.ok(drawing.Drawing.spPr.Fill.fill instanceof AscFormat.CGradFill, "Shape created and filled with gradient");
        assert.strictEqual(drawing.Drawing.spPr.Fill.fill.colors.length, 2, 'Check colors of gradient amount');

        let firstColor = drawing.Drawing.spPr.Fill.fill.colors[0].color.color.RGBA;

        assert.strictEqual(firstColor.R, 255, 'Check color of first gradient fill R');
        assert.strictEqual(firstColor.G, 213, 'Check color of first gradient fill G');
        assert.strictEqual(firstColor.B, 191, 'Check color of first gradient fill B');

        let secondColor = drawing.Drawing.spPr.Fill.fill.colors[1].color.color.RGBA;

        assert.strictEqual(secondColor.R, 255, 'Check color of second gradient fill R');
        assert.strictEqual(secondColor.G, 111, 'Check color of second gradient fill G');
        assert.strictEqual(secondColor.B, 61,  'Check color of second gradient fill B');

	});

	QUnit.test("Test: SetOutLine", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);

		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(100, 100, 100));
		const initialStroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("rect", 200 * 36000, 100 * 36000, fill, initialStroke);

		drawing.SetPosition(500000, 500000);
		slide.AddObject(drawing);

		assert.strictEqual(drawing.Drawing.spPr.ln.w, null, 'Initial outline width should be null');
		const strokeFill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(255, 0, 0));
		const newStroke = AscTest.JsApi.CreateStroke(25400, strokeFill); // 1pt = 12700, so 2pt = 25400

		const result = drawing.SetOutLine(newStroke);

		assert.strictEqual(result, true, 'SetOutLine should return true on success');
		assert.strictEqual(drawing.Drawing.spPr.ln.w, 25400, 'Outline width should be 25400 (2pt)');

		const outlineColor = drawing.Drawing.spPr.ln.Fill.fill.color.color.RGBA;
		assert.strictEqual(outlineColor.R, 255, 'Outline color R should be 255');
		assert.strictEqual(outlineColor.G, 0, 'Outline color G should be 0');
		assert.strictEqual(outlineColor.B, 0, 'Outline color B should be 0');

		const invalidResult = drawing.SetOutLine(null);
		assert.strictEqual(invalidResult, false, 'SetOutLine should return false with null parameter');

		const invalidResult2 = drawing.SetOutLine({});
		assert.strictEqual(invalidResult2, false, 'SetOutLine should return false with invalid parameter');
	});

	QUnit.test("Test: GetName", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		const name = drawing.GetName();
		assert.strictEqual(typeof name, 'string', 'Check GetName returns a string');
		assert.ok(name.length > 0, 'Check drawing name is not empty');
	});

	QUnit.test("Test: SetName", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		let result = drawing.SetName("TestShape");
		assert.strictEqual(result, true, 'Check SetName returns true');
		assert.strictEqual(drawing.GetName(), "TestShape", 'Check drawing name is set correctly');

		let result2 = drawing.SetName("");
		assert.strictEqual(result2, false, 'Check SetName returns false for empty string');

		let result3 = drawing.SetName(null);
		assert.strictEqual(result3, false, 'Check SetName returns false for null');

		let result4 = drawing.SetName(undefined);
		assert.strictEqual(result4, false, 'Check SetName returns false for undefined');

		// Test that setting duplicate name causes previous shape to get default name
		const drawing2 = AscTest.JsApi.CreateShape("rect", 150 * 36000, 80 * 36000, fill, stroke);
		drawing2.SetPosition(608400, 2267200);
		slide.AddObject(drawing2);

		drawing.SetName("DuplicateName");
		const firstDrawingName = drawing.GetName();
		assert.strictEqual(firstDrawingName, "DuplicateName", 'Check first drawing has duplicate name');

		drawing2.SetName("DuplicateName");

		assert.strictEqual(drawing2.GetName(), "DuplicateName", 'Check second drawing has the duplicate name');
		assert.notStrictEqual(drawing.GetName(), "DuplicateName", 'Check first drawing name changed from duplicate');
		assert.notStrictEqual(drawing.GetName(), firstDrawingName, 'Check first drawing has a new default name');
	});

	QUnit.test("Test: Select", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		drawing.Select();
		assert.ok(true, 'Check Select method works');
		assert.ok(drawing.Drawing.getDrawingObjectsController().selectedObjects.includes(drawing.Drawing), 'Check drawing is selected in presentation');

		drawing.Select(true);
		assert.ok(true, 'Check Select with isReplace=true works');
	});

	QUnit.test("Test: Unselect", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		drawing.Select();
		assert.ok(drawing.Drawing.getDrawingObjectsController().selectedObjects.includes(drawing.Drawing), 'Check drawing is selected before unselect');

		let result = drawing.Unselect();
		assert.strictEqual(result, true, 'Check Unselect returns true');
		assert.ok(!drawing.Drawing.getDrawingObjectsController().selectedObjects.includes(drawing.Drawing), 'Check drawing is not selected after unselect');
	});

	QUnit.test("Test: GetFlipH", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(255, 111, 61));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		assert.strictEqual(drawing.GetFlipH(), false, 'Check drawing horizontal flip === false');
		drawing.SetFlipH(true);
		assert.strictEqual(drawing.GetFlipH(), true, 'Check drawing horizontal flip === true');
	});

	QUnit.test("Test: GetFlipV", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(255, 111, 61));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		assert.strictEqual(drawing.GetFlipV(), false, 'Check drawing vertical flip === false');
		drawing.SetFlipV(true);
		assert.strictEqual(drawing.GetFlipV(), true, 'Check drawing vertical flip === true');
	});

	QUnit.test("Test: SetFlipH", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(255, 111, 61));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		let result = drawing.SetFlipH(true);
		assert.strictEqual(result, true, 'Check SetFlipH returns true');
		assert.strictEqual(drawing.GetFlipH(), true, 'Check drawing horizontal flip === true after SetFlipH');

		result = drawing.SetFlipH(false);
		assert.strictEqual(result, true, 'Check SetFlipH returns true');
		assert.strictEqual(drawing.GetFlipH(), false, 'Check drawing horizontal flip === false after SetFlipH');

		result = drawing.SetFlipH("invalid");
		assert.strictEqual(result, false, 'Check SetFlipH returns false for invalid parameter');
	});

	QUnit.test("Test: SetFlipV", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(255, 111, 61));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const drawing = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);

		drawing.SetPosition(608400, 1267200);
		slide.AddObject(drawing);

		let result = drawing.SetFlipV(true);
		assert.strictEqual(result, true, 'Check SetFlipV returns true');
		assert.strictEqual(drawing.GetFlipV(), true, 'Check drawing vertical flip === true after SetFlipV');

		result = drawing.SetFlipV(false);
		assert.strictEqual(result, true, 'Check SetFlipV returns true');
		assert.strictEqual(drawing.GetFlipV(), false, 'Check drawing vertical flip === false after SetFlipV');

		result = drawing.SetFlipV("invalid");
		assert.strictEqual(result, false, 'Check SetFlipV returns false for invalid parameter');
	});
});

