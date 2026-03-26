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

	QUnit.test("SetVisible", function (assert) {
		const sheet = AscTest.JsApi.AddSheet("TestSetVisible_" + Date.now());

		assert.strictEqual(sheet.SetVisible(false), true, "SetVisible returns true");
		assert.strictEqual(sheet.GetVisible(), false, "Sheet is hidden after SetVisible(false)");

		assert.strictEqual(sheet.SetVisible(true), true, "SetVisible returns true");
		assert.strictEqual(sheet.GetVisible(), true, "Sheet is visible after SetVisible(true)");

		sheet.Delete();
	});

	QUnit.test("SetActive", function (assert) {
		const newSheet = AscTest.JsApi.AddSheet("TestSetActive_" + Date.now());

		const firstSheet = AscTest.JsApi.GetSheet(0);
		assert.strictEqual(firstSheet.SetActive(), true, "SetActive returns true");
		assert.strictEqual(AscTest.JsApi.GetActiveSheet().GetName(), firstSheet.GetName(), "First sheet is active after SetActive");

		newSheet.Delete();
	});

	QUnit.test("SetRowHeight", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();
		assert.strictEqual(sheet.SetRowHeight(0, 30), true, "SetRowHeight returns true");
	});

	QUnit.test("SetDisplayGridlines", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();

		assert.strictEqual(sheet.SetDisplayGridlines(false), true, "SetDisplayGridlines returns true");
		assert.strictEqual(sheet.worksheet.sheetViews[0].showGridLines, false, "Gridlines are hidden after SetDisplayGridlines(false)");

		assert.strictEqual(sheet.SetDisplayGridlines(true), true, "SetDisplayGridlines returns true");
		assert.strictEqual(sheet.worksheet.sheetViews[0].showGridLines, true, "Gridlines are visible after SetDisplayGridlines(true)");
	});

	QUnit.test("SetDisplayHeadings", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();

		assert.strictEqual(sheet.SetDisplayHeadings(false), true, "SetDisplayHeadings returns true");
		assert.strictEqual(sheet.worksheet.sheetViews[0].showRowColHeaders, false, "Headings are hidden after SetDisplayHeadings(false)");

		assert.strictEqual(sheet.SetDisplayHeadings(true), true, "SetDisplayHeadings returns true");
		assert.strictEqual(sheet.worksheet.sheetViews[0].showRowColHeaders, true, "Headings are visible after SetDisplayHeadings(true)");
	});

	QUnit.test("SetLeftMargin", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();
		const nPoints = 72;

		assert.strictEqual(sheet.SetLeftMargin(nPoints), true, "SetLeftMargin returns true");
		assert.strictEqual(sheet.GetLeftMargin(), nPoints, "GetLeftMargin returns the value that was set");
	});

	QUnit.test("SetRightMargin", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();
		const nPoints = 72;

		assert.strictEqual(sheet.SetRightMargin(nPoints), true, "SetRightMargin returns true");
		assert.strictEqual(sheet.GetRightMargin(), nPoints, "GetRightMargin returns the value that was set");
	});

	QUnit.test("SetTopMargin", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();
		const nPoints = 72;

		assert.strictEqual(sheet.SetTopMargin(nPoints), true, "SetTopMargin returns true");
		assert.strictEqual(sheet.GetTopMargin(), nPoints, "GetTopMargin returns the value that was set");
	});

	QUnit.test("SetBottomMargin", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();
		const nPoints = 72;

		assert.strictEqual(sheet.SetBottomMargin(nPoints), true, "SetBottomMargin returns true");
		assert.strictEqual(sheet.GetBottomMargin(), nPoints, "GetBottomMargin returns the value that was set");
	});

	QUnit.test("SetPageOrientation", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();

		assert.strictEqual(sheet.SetPageOrientation("xlLandscape"), true, "SetPageOrientation returns true");
		assert.strictEqual(sheet.GetPageOrientation(), "xlLandscape", "GetPageOrientation returns xlLandscape after setting it");

		assert.strictEqual(sheet.SetPageOrientation("xlPortrait"), true, "SetPageOrientation returns true");
		assert.strictEqual(sheet.GetPageOrientation(), "xlPortrait", "GetPageOrientation returns xlPortrait after setting it");
	});

	QUnit.test("SetPrintHeadings", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();

		assert.strictEqual(sheet.SetPrintHeadings(true), true, "SetPrintHeadings returns true");
		assert.strictEqual(sheet.GetPrintHeadings(), true, "GetPrintHeadings returns true after SetPrintHeadings(true)");

		assert.strictEqual(sheet.SetPrintHeadings(false), true, "SetPrintHeadings returns true");
		assert.strictEqual(sheet.GetPrintHeadings(), false, "GetPrintHeadings returns false after SetPrintHeadings(false)");
	});

	QUnit.test("SetPrintGridlines", function (assert) {
		const sheet = AscTest.JsApi.GetActiveSheet();

		assert.strictEqual(sheet.SetPrintGridlines(true), true, "SetPrintGridlines returns true");
		assert.strictEqual(sheet.GetPrintGridlines(), true, "GetPrintGridlines returns true after SetPrintGridlines(true)");

		assert.strictEqual(sheet.SetPrintGridlines(false), true, "SetPrintGridlines returns true");
		assert.strictEqual(sheet.GetPrintGridlines(), false, "GetPrintGridlines returns false after SetPrintGridlines(false)");
	});

	QUnit.test("Delete", function (assert) {
		const sheetName = "TestDelete_" + Date.now();
		const sheet = AscTest.JsApi.AddSheet(sheetName);
		const countBefore = AscTest.JsApi.GetSheets().length;

		assert.strictEqual(sheet.Delete(), true, "Delete returns true");
		assert.strictEqual(AscTest.JsApi.GetSheets().length, countBefore - 1, "Sheet count decreased after Delete");
		assert.strictEqual(AscTest.JsApi.GetSheet(sheetName), null, "Deleted sheet is no longer accessible by name");
	});

})(window);

