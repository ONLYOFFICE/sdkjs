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

"use strict";

(function (window)
{
	QUnit.module("ApiDrawing");
	QUnit.test("Fill", function (assert) {
		let worksheet = AscTest.JsApi.GetActiveSheet()

		let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());

		let shape = worksheet.AddShape("ellipse", 50 * 36000, 50 * 36000, fill, stroke, 0, 0, 0, 0);
		assert.ok(true, 'Add new ellipse shape');
		
		let gs1 = AscTest.JsApi.CreateGradientStop(AscTest.JsApi.CreateRGBColor(255, 213, 191), 0);
		let gs2 = AscTest.JsApi.CreateGradientStop(AscTest.JsApi.CreateRGBColor(255, 111, 61), 100000);
		fill = AscTest.JsApi.CreateRadialGradientFill([gs1, gs2]);
		shape.Fill(fill);

        assert.ok(shape.Drawing.spPr.Fill.fill instanceof AscFormat.CGradFill, "Shape created and filled with gradient");
        assert.strictEqual(shape.Drawing.spPr.Fill.fill.colors.length, 2, 'Check colors of gradient amount');

        let firstColor = shape.Drawing.spPr.Fill.fill.colors[0].color.color.RGBA;

        assert.strictEqual(firstColor.R, 255, 'Check color of first gradient fill R');
        assert.strictEqual(firstColor.G, 213, 'Check color of first gradient fill G');
        assert.strictEqual(firstColor.B, 191, 'Check color of first gradient fill B');

        let secondColor = shape.Drawing.spPr.Fill.fill.colors[1].color.color.RGBA;

        assert.strictEqual(secondColor.R, 255, 'Check color of second gradient fill R');
        assert.strictEqual(secondColor.G, 111, 'Check color of second gradient fill G');
        assert.strictEqual(secondColor.B, 61,  'Check color of second gradient fill B');
	});

	QUnit.test("SetOutLine", function (assert) {
		let worksheet = AscTest.JsApi.GetActiveSheet()

		let fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		let stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());

		let shape = worksheet.AddShape("ellipse", 50 * 36000, 50 * 36000, fill, stroke, 0, 0, 0, 0);
		assert.ok(true, 'Add new ellipse shape');

		let outlineFill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(255, 111, 61));
		let outline = AscTest.JsApi.CreateStroke(1 * 36000, outlineFill);
		shape.SetOutLine(outline);

		assert.ok(shape.Drawing.spPr.ln, "Shape outline is defined");
		assert.ok(shape.Drawing.spPr.ln.Fill.fill instanceof AscFormat.CSolidFill, "Shape outline filled with solid fill");

		let outlineColor = shape.Drawing.spPr.ln.Fill.fill.color.color.RGBA;

		assert.strictEqual(outlineColor.R, 255, 'Check color of outline R');
		assert.strictEqual(outlineColor.G, 111, 'Check color of outline G');
		assert.strictEqual(outlineColor.B, 61, 'Check color of outline B');
		assert.strictEqual(shape.Drawing.spPr.ln.w, 1 * 36000, 'Check outline width');
	});

})(window);
