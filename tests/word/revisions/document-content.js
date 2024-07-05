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
	
	QUnit.test("Remove/replace text in a single run", function (assert)
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
	
});
