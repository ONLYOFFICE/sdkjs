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
	QUnit.module("ApiWorksheet");
	QUnit.test("GetSelectedShapes", function (assert) {
		let worksheet = AscTest.JsApi.GetActiveSheet();
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());

		let shapes = [];
		// Create 3 shapes, but select only 2 of them (skip the middle one)
		for(let nShape = 0; nShape < 3; nShape++)
		{
			let shape = worksheet.AddShape("ellipse", 50 * 36000, 50 * 36000, fill, stroke, 0, 0, 0, 0);
			assert.ok(shape, 'Ellipse shape created');
			shapes.push(shape);
			if (nShape !== 1) {
				shape.Select();
			}
		}

		let selectedShapes = worksheet.GetSelectedShapes();
		assert.strictEqual(
			selectedShapes.length,
			2,
			"Count of selected shapes is 2 (first and third)"
		);

		// Verify that selected shapes are the ones we actually selected
		let selectedDrawings = selectedShapes.map(s => s.Drawing);
		assert.ok(
			selectedDrawings.includes(shapes[0].Drawing),
			"First shape is selected"
		);
		assert.ok(
			selectedDrawings.includes(shapes[2].Drawing),
			"Third shape is selected"
		);
		assert.ok(
			!selectedDrawings.includes(shapes[1].Drawing),
			"Second shape is not selected"
		);
	});
	
	QUnit.test("GetSelectedDrawings", function (assert) {
		let worksheet = AscTest.JsApi.GetActiveSheet();

		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());

		let shapes = [];
		// Create 3 shapes, but select only 2 of them (skip the middle one)
		for(let nShape = 0; nShape < 3; nShape++)
		{
			let shape = worksheet.AddShape("ellipse", 50 * 36000, 50 * 36000, fill, stroke, 0, 0, 0, 0);
			assert.ok(shape, 'Ellipse shape created');
			shapes.push(shape);
			if (nShape !== 1) {
				shape.Select();
			}
		}

		// Add first image and select it
		let image1 = worksheet.AddImage("https://static.onlyoffice.com/assets/docs/samples/img/presentation_sky.png", 60 * 36000, 35 * 36000, 0, 2 * 36000, 0, 3 * 36000);
		assert.ok(image1, 'First image created');
		image1.Select();

		// Add second image but don't select it
		let image2 = worksheet.AddImage("https://static.onlyoffice.com/assets/docs/samples/img/presentation_sky.png", 60 * 36000, 35 * 36000, 0, 2 * 36000, 0, 3 * 36000);
		assert.ok(image2, 'Second image created');

		let selectedDrawings = worksheet.GetSelectedDrawings();
		assert.strictEqual(
			selectedDrawings.length,
			3,
			"Count of selected drawings is 3 (2 shapes + 1 image)"
		);

		// Verify that selected drawings are the ones we actually selected
		let selectedDrawingsInner = selectedDrawings.map(d => d.Drawing);
		assert.ok(
			selectedDrawingsInner.includes(shapes[0].Drawing),
			"First shape is selected"
		);
		assert.ok(
			selectedDrawingsInner.includes(shapes[2].Drawing),
			"Third shape is selected"
		);
		assert.ok(
			selectedDrawingsInner.includes(image1.Drawing),
			"First image is selected"
		);
		assert.ok(
			!selectedDrawingsInner.includes(shapes[1].Drawing),
			"Second shape is not selected"
		);
		assert.ok(
			!selectedDrawingsInner.includes(image2.Drawing),
			"Second image is not selected"
		);
	});

	QUnit.test("GetDrawingsByName", function (assert) {
		let worksheet = AscTest.JsApi.GetActiveSheet();
		let workbook = AscTest.JsApi.GetActiveWorkbook();

		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(255, 111, 61));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());

		let shape1 = worksheet.AddShape("cube", 3212465, 963295, fill, stroke, 0, 0, 0, 0);
		let shape2 = worksheet.AddShape("rect", 3212465, 963295, fill, stroke, 0, 0, 0, 0);

		shape1.SetName("Shape1");
		shape2.SetName("Shape2");

		let drawings = workbook.GetDrawingsByName(["Shape1", "Shape2"]);
		assert.strictEqual(drawings.length, 2, 'Check GetDrawingsByName returns 2 drawings');

		let drawingsFiltered = workbook.GetDrawingsByName(["Shape1"]);
		assert.strictEqual(drawingsFiltered.length, 1, 'Check GetDrawingsByName returns 1 drawing');
		assert.strictEqual(drawingsFiltered[0].GetName(), "Shape1", 'Check filtered drawing has correct name');
	});

})(window);

