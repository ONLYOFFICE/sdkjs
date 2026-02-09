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
	QUnit.module("ApiShape");

    function CreateSlide()
	{
		logicDocument.addNextSlide(0);
		editor.WordControl.Thumbnails.CalculatePlaces();
	}

	QUnit.test("Test: SetPaddings", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);

		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(200, 200, 200));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const shape = AscTest.JsApi.CreateShape("rect", 300 * 36000, 150 * 36000, fill, stroke);

		slide.AddObject(shape);

		const docContent = shape.GetDocContent();
		const paragraph = AscTest.JsApi.CreateParagraph();
		paragraph.AddText("Test text with paddings");
		docContent.Push(paragraph);

		const result = shape.SetPaddings(4 * 36000, 2 * 36000, 3 * 36000, 5 * 36000);

		assert.strictEqual(result, true, 'SetPaddings should return true');

		const bodyPr = shape.Shape.getBodyPr();

		assert.ok(bodyPr.lIns === 4, 'Left padding should be set 4');
		assert.ok(bodyPr.tIns === 2, 'Top padding should be set 2');
		assert.ok(bodyPr.rIns === 3, 'Right padding should be set 3');
		assert.ok(bodyPr.bIns === 5, 'Bottom padding should be set 5');
	});
});
