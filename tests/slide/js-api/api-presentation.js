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
	QUnit.module("ApiPresentation");

    function CreateSlide()
	{
		logicDocument.addNextSlide(0);
		editor.WordControl.Thumbnails.CalculatePlaces();
	}

	QUnit.test("Test: GetDrawingsByName", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);
		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(51, 51, 51));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());

		const drawing1 = AscTest.JsApi.CreateShape("cube", 150 * 36000, 80 * 36000, fill, stroke);
		drawing1.SetPosition(608400, 1267200);
		slide.AddObject(drawing1);

		const drawing2 = AscTest.JsApi.CreateShape("rect", 150 * 36000, 80 * 36000, fill, stroke);
		drawing2.SetPosition(608400, 2267200);
		slide.AddObject(drawing2);

		drawing1.SetName("Shape1");
		drawing2.SetName("Shape2");

		let drawings = presentation.GetDrawingsByName(["Shape1", "Shape2"]);
		assert.strictEqual(drawings.length, 2, 'Check GetDrawingsByName returns 2 drawings');

		let drawingsFiltered = presentation.GetDrawingsByName(["Shape1"]);
		assert.strictEqual(drawingsFiltered.length, 1, 'Check GetDrawingsByName returns 1 drawing');
		assert.strictEqual(drawingsFiltered[0].GetName(), "Shape1", 'Check filtered drawing has correct name');
	});
});

