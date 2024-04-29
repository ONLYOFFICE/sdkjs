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

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();
	logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);
	
	QUnit.module("Test the revisions in a paragraph");
	
	function fillDocument(runData)
	{
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		
		runData.forEach(function(d)
		{
			let run = AscTest.CreateRun();
			run.AddText(d.text);
			if (undefined !== d.reviewType)
				run.SetReviewType(d.reviewType);
			else
				run.SetReviewType(reviewtype_Common);
			
			p.AddToContentToEnd(run);
		});
		
		return p;
	}
	
	function fillDocument_1234test()
	{
		return fillDocument([
			{text : "1234", reviewType : reviewtype_Add},
			{text : "test"},
		]);
	}
	
	QUnit.test("Remove/replace text in a single run", function (assert)
	{
		AscTest.SetTrackRevisions(true);
		let p = fillDocument_1234test();
		assert.strictEqual(AscTest.GetParagraphText(p), "1234test", 'Check paragraph text');
		
		// TODO: We only checked appearance of the text, but didn't check review mode of the added text
		AscTest.SelectParagraphRange(p, 1, 3);
		AscTest.EnterText("QQQ");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1QQQ4"],
				[reviewtype_Common, "test"],
			],
			"Select text. Enter text over selection"
		);
		
		p = fillDocument_1234test();
		AscTest.SelectParagraphRange(p, 1, 3);
		AscTest.PressKey(AscTest.Key.delete);
		AscTest.EnterText("QQQ");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1QQQ4"],
				[reviewtype_Common, "test"],
			],
			"Select text. Press delete button. Enter text over selection"
		);
		
		p = fillDocument_1234test();
		AscTest.SelectParagraphRange(p, 1, 3);
		AscTest.PressKey(AscTest.Key.backspace);
		AscTest.EnterText("QQQ");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1QQQ4"],
				[reviewtype_Common, "test"],
			],
			"Select text. Press backspace button. Enter text"
		);
	});
	
	QUnit.test("Remove/replace text in several runs", function (assert)
	{
		AscTest.SetTrackRevisions(true);
		let p = fillDocument_1234test();
		assert.strictEqual(AscTest.GetParagraphText(p), "1234test", 'Check paragraph text');

		// TODO: We only checked appearance of the text, but didn't check review mode of the added text
		AscTest.SelectParagraphRange(p, 1, 6);
		AscTest.EnterText("ABC");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1ABC"],
				[reviewtype_Remove, "te"],
				[reviewtype_Common, "st"],
			],
			"Select text and enter text"
		);

		p = fillDocument_1234test();
		AscTest.SelectParagraphRange(p, 1, 6);
		AscTest.PressKey(AscTest.Key.delete);
		AscTest.EnterText("ABC");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1"],
				[reviewtype_Remove, "te"],
				[reviewtype_Add, "ABC"],
				[reviewtype_Common, "st"],
			],
			"Select text. Press delete button. Enter text"
		);

		p = fillDocument_1234test();
		AscTest.SelectParagraphRange(p, 1, 6);
		AscTest.PressKey(AscTest.Key.backspace);
		AscTest.EnterText("ABC");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1ABC"],
				[reviewtype_Remove, "te"],
				[reviewtype_Common, "st"],
			],
			"Select text. Press backspace button. Enter text"
		);

		p = fillDocument([
			{text : "Before"},
			{text : "1234", reviewType : reviewtype_Add},
			{text : "after"},
		]);
		AscTest.SelectParagraphRange(p, 8, 12);
		AscTest.EnterText("777");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "Before"],
				[reviewtype_Add, "12777"],
				[reviewtype_Remove, "af"],
				[reviewtype_Common, "ter"],
			],
			"Select text and enter text"
		);

		p = fillDocument([
			{text : "Before"},
			{text : "1234", reviewType : reviewtype_Add},
			{text : "after"},
		]);
		AscTest.SelectParagraphRange(p, 8, 12);
		AscTest.PressKey(AscTest.Key.delete);
		AscTest.EnterText("777");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "Before"],
				[reviewtype_Add, "12"],
				[reviewtype_Remove, "af"],
				[reviewtype_Add, "777"],
				[reviewtype_Common, "ter"],
			],
			"Select text. Press delete button. Enter text"
		);
		
		p = fillDocument([
			{text : "Before"},
			{text : "1234", reviewType : reviewtype_Add},
			{text : "after"},
		]);
		AscTest.SelectParagraphRange(p, 8, 12);
		AscTest.PressKey(AscTest.Key.backspace);
		AscTest.EnterText("777");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "Before"],
				[reviewtype_Add, "12777"],
				[reviewtype_Remove, "af"],
				[reviewtype_Common, "ter"],
			],
			"Select text. Press backspace button. Enter text"
		);
	});
});
