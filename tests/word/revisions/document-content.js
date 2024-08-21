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
	
	QUnit.module("Test the revisions on the document content level");
	
	AscTest.PressEnter = function()
	{
		logicDocument.AddNewParagraph();
	};
	
	QUnit.test("Test adding a new paragraph if track revisions is on", function (assert)
	{
		let isTrack = false;
		
		let p1, p2, p3, p4;
		function initTestDocument()
		{
			AscTest.SetTrackRevisions(false);
			AscTest.ClearDocument();
			
			p1 = AscTest.CreateParagraph();
			p2 = AscTest.CreateParagraph();
			p3 = AscTest.CreateParagraph();
			p4 = AscTest.CreateParagraph();
			logicDocument.PushToContent(p1);
			logicDocument.PushToContent(p2);
			logicDocument.PushToContent(p3);
			logicDocument.PushToContent(p4);
			
			AscTest.MoveCursorToParagraph(p1);
			AscTest.EnterText("Test");
			AscTest.MoveCursorToParagraph(p2);
			AscTest.EnterText("Text text");
			AscTest.MoveCursorToParagraph(p3);
			AscTest.EnterText("Test text");
			
			p2.SetReviewType(reviewtype_Add);
			p3.SetReviewType(reviewtype_Remove);
			AscTest.SetTrackRevisions(isTrack);
		}
		
		function testReviewTypes(types)
		{
			assert.strictEqual(logicDocument.GetElementsCount(), types.length, "Check number of paragraphs");
			if (types.length !== logicDocument.GetElementsCount())
				return;
			
			for (let i = 0; i < types.length; ++i)
			{
				assert.strictEqual(logicDocument.GetElement(i).GetReviewType(), types[i], "Check review type for " + (i) + " paragraph");
			}
		}
		
		function testAddingNewParagraph(paraIndex, types)
		{
			// Предполагаем, что параграф не пустой (как минимум 3 символа), проверяем добавление параграфа
			// когда курсор стоит в начале/середине/конце
			initTestDocument();
			let p = logicDocument.GetElement(paraIndex);
			AscTest.MoveCursorToParagraph(p, true);
			AscTest.PressEnter();
			testReviewTypes(types);
			
			initTestDocument();
			p = logicDocument.GetElement(paraIndex)
			AscTest.MoveCursorToParagraph(p, true);
			AscTest.MoveCursorRight(false, false, 2);
			AscTest.PressEnter();
			testReviewTypes(types);
			
			initTestDocument();
			p = logicDocument.GetElement(paraIndex)
			AscTest.MoveCursorToParagraph(p, false);
			AscTest.PressEnter();
			testReviewTypes(types);
		}
		
		
		initTestDocument();
		testReviewTypes([
			reviewtype_Common,
			reviewtype_Add,
			reviewtype_Remove,
			reviewtype_Common
		]);
		
		AscTest.SetTrackRevisions(false);
		
		testAddingNewParagraph(0,
			[
				reviewtype_Common,
				reviewtype_Common,
				reviewtype_Add,
				reviewtype_Remove,
				reviewtype_Common
			]
		);

		testAddingNewParagraph(1,
			[
				reviewtype_Common,
				reviewtype_Common,
				reviewtype_Add,
				reviewtype_Remove,
				reviewtype_Common
			]
		);

		testAddingNewParagraph(2,
			[
				reviewtype_Common,
				reviewtype_Add,
				reviewtype_Common,
				reviewtype_Remove,
				reviewtype_Common
			]
		);
		
		AscTest.SetTrackRevisions(true);
		isTrack = true;
		
		testAddingNewParagraph(0,
			[
				reviewtype_Add,
				reviewtype_Common,
				reviewtype_Add,
				reviewtype_Remove,
				reviewtype_Common
			]
		);
		
		testAddingNewParagraph(1,
			[
				reviewtype_Common,
				reviewtype_Add,
				reviewtype_Add,
				reviewtype_Remove,
				reviewtype_Common
			]
		);

		testAddingNewParagraph(2,
			[
				reviewtype_Common,
				reviewtype_Add,
				reviewtype_Add,
				reviewtype_Remove,
				reviewtype_Common
			]
		);
		
		AscTest.SetTrackRevisions(false);
	});
	
	QUnit.testStart(function()
	{
		AscTest.ClearDocument();
		AscTest.SetTrackRevisions(false);
	});
	
	QUnit.module("Remove/replace text in a block-level sdt");
	QUnit.test("Check replacing text in a block-level content control (bug 67071)", function(assert)
	{
		let p = AscTest.CreateParagraph();
		let cc = AscTest.CreateBlockLvlSdt();
		
		logicDocument.AddToContent(0, p);
		logicDocument.AddToContent(0, cc);
		
		p = cc.GetElement(0);
		
		cc.SelectContentControl();
		AscTest.EnterText("Text");
		
		cc.SelectContentControl();
		
		AscTest.SetTrackRevisions(true);
		
		AscTest.EnterText("123");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Remove, "Text"],
				[reviewtype_Add, "123"],
			],
			"Select text. Enter text over selection"
		);
		
		AscTest.AcceptAllRevisionChanges();
		assert.strictEqual(cc.IsUseInDocument(), true, "Check if content control is still present in the document");
		assert.deepEqual(
			AscTest.GetBlockLevelSdtReviewText(cc),
			[
				[reviewtype_Common, "123"]
			],
			"Check content control text after accepting all changes"
		);
	});
	QUnit.test("Check accepting all changes when entire content of a block-level sdt was deleted", function(assert)
	{
		let p = AscTest.CreateParagraph();
		let cc = AscTest.CreateBlockLvlSdt();

		logicDocument.AddToContent(0, p);
		logicDocument.AddToContent(0, cc);

		p = cc.GetElement(0);

		cc.SelectContentControl();
		AscTest.EnterText("Text");

		AscTest.SetTrackRevisions(true);

		cc.SelectContentControl();
		logicDocument.Remove(-1);

		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Remove, "Text"],
			],
			"Select text. Enter text over selection"
		);

		AscTest.AcceptAllRevisionChanges();
		assert.strictEqual(cc.IsUseInDocument(), false, "Check if content control is still present in the document");
		assert.strictEqual(logicDocument.GetElementsCount(), 1, "Check the number of elements in the document");
	});
	QUnit.test("Check accepting changes in the special case when entire document was deleted (including a block-level sdt) (bug 69615)", function(assert)
	{
		function test(bySelection)
		{
			AscTest.ClearDocument();
			AscTest.SetTrackRevisions(false);
			
			let pAfter  = AscTest.CreateParagraph();
			let pBefore = AscTest.CreateParagraph();
			let cc      = AscTest.CreateBlockLvlSdt();
			
			logicDocument.AddToContent(0, pAfter);
			logicDocument.AddToContent(0, cc);
			logicDocument.AddToContent(0, pBefore);
			
			AscTest.AddTextToParagraph(pAfter, "After");
			AscTest.AddTextToParagraph(pBefore, "Before");
			
			p = cc.GetElement(0);
			
			cc.SelectContentControl();
			AscTest.EnterText("Inside content control");
			
			AscTest.SetTrackRevisions(true);
			
			logicDocument.SelectAll();
			logicDocument.Remove(-1);
			
			if (bySelection)
			{
				logicDocument.SelectAll();
				AscTest.AcceptRevisionChangesBySelection();
			}
			else
			{
				AscTest.AcceptAllRevisionChanges();
			}
			
			assert.true(true, bySelection ? "BySelection" : "All");
			
			assert.strictEqual(cc.IsUseInDocument(), false, "Check if content control is still present in the document");
			assert.strictEqual(logicDocument.GetElementsCount(), 1, "Check the number of elements in the document");
			assert.strictEqual(logicDocument.GetElement(0).IsParagraph() && logicDocument.GetElement(0).IsEmpty(), true, "Check if the last element in the document is an empty paragraph");
		}
		
		test(false);
		test(true);
	});
	QUnit.test("Check rejecting changes in the special case when entire document was added (including a block-level sdt)", function(assert)
	{
		function test(bySelection)
		{
			AscTest.ClearDocument();
			AscTest.SetTrackRevisions(true);
			
			let pAfter  = AscTest.CreateParagraph();
			let pBefore = AscTest.CreateParagraph();
			let cc      = AscTest.CreateBlockLvlSdt();
			
			logicDocument.AddToContent(0, pAfter);
			logicDocument.AddToContent(0, cc);
			logicDocument.AddToContent(0, pBefore);
			
			AscTest.AddTextToParagraph(pAfter, "After");
			AscTest.AddTextToParagraph(pBefore, "Before");
			
			p = cc.GetElement(0);
			
			cc.SelectContentControl();
			AscTest.EnterText("Inside content control");
			
			logicDocument.SelectAll();
			
			if (bySelection)
				AscTest.RejectRevisionChangesBySelection();
			else
				AscTest.RejectAllRevisionChanges();
			
			assert.true(true, bySelection ? "BySelection" : "All");
			
			assert.strictEqual(cc.IsUseInDocument(), false, "Check if content control is still present in the document");
			assert.strictEqual(logicDocument.GetElementsCount(), 1, "Check the number of elements in the document");
			assert.strictEqual(logicDocument.GetElement(0).IsParagraph() && logicDocument.GetElement(0).IsEmpty(), true, "Check if the last element in the document is an empty paragraph");
		}
		
		test(false);
		test(true);
	});
	
});
