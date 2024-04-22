/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument()
	
	QUnit.test("Test various actions with checkbox content control", function (assert)
	{
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		AscTest.MoveCursorToParagraph(p);
		
		let checkboxPr = new AscWord.CSdtCheckBoxPr();
		checkboxPr.SetCheckedSymbol("T".codePointAt(0));
		checkboxPr.SetUncheckedSymbol("F".codePointAt(0));
		let checkbox = logicDocument.AddContentControlCheckBox(checkboxPr);

		assert.strictEqual(checkbox.IsUseInDocument(), true, "Check if checkbox is added to the document");
		assert.strictEqual(checkbox.IsCheckBoxChecked(), false);
		assert.strictEqual(AscTest.GetParagraphText(p), "F");
		
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), true);
		assert.strictEqual(AscTest.GetParagraphText(p), "T", "Check toggle checkbox in normal mode");
		
		AscTest.SetTrackRevisions(true);
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), false);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "F"],
				[reviewtype_Remove, "T"],
			],
			"Check toggle checkbox in review"
		);
		
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), true);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "T"],
			],
			"Check toggle checkbox in review"
		);
		AscTest.SetTrackRevisions(false);
		
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), false);
		assert.strictEqual(AscTest.GetParagraphText(p), "F", "Check toggle checkbox in normal mode");
		
		AscTest.SetTrackRevisions(true);
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), true);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "T"],
				[reviewtype_Remove, "F"],
			],
			"Check toggle checkbox in review"
		);
		
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), false);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "F"],
			],
			"Check toggle checkbox in review"
		);
		
		
		let p2 = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p2);
		AscTest.MoveCursorToParagraph(p2);
		
		let checkbox2 = logicDocument.AddContentControlCheckBox(checkboxPr);
		assert.strictEqual(checkbox2.IsUseInDocument(), true, "Check if checkbox is added to the document");
		assert.strictEqual(checkbox2.IsCheckBoxChecked(), false);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p2),
			[
				[reviewtype_Add, "F"],
			],
			"Check adding a checkbox in review"
		);
		
		checkbox2.ToggleCheckBox();
		assert.strictEqual(checkbox2.IsCheckBoxChecked(), true);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p2),
			[
				[reviewtype_Add, "T"]
			],
			"Check toggle checkbox in review (check box was added in review)"
		);
		
		checkbox2.ToggleCheckBox();
		assert.strictEqual(checkbox2.IsCheckBoxChecked(), false);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p2),
			[
				[reviewtype_Add, "F"]
			],
			"Check toggle checkbox in review (check box was added in review)"
		);
		
		
		AscTest.SetTrackRevisions(false);
		
	});
	
});
