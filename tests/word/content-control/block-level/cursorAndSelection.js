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

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument()
	
	function CreateContentControl()
	{
		let cc = new AscWord.CBlockLevelSdt();
		cc.SetPlaceholder(c_oAscDefaultPlaceholderName.Text);
		cc.ReplacePlaceHolderWithContent();
		cc.SetShowingPlcHdr(false);
		return cc;
	}

	function CreateParagraphWithText(text)
	{
		let p = AscTest.CreateParagraph();
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		run.AddText(text);
		return p;
	}
	
	QUnit.module("Test the positioning of the cursor and selection for inline-level content controls");
	
	QUnit.test("Test remove/delete after/before content control", function (assert)
	{
		AscTest.SetTrackRevisions(false);
		AscTest.ClearDocument();

		function CreateFilledContentControl(texts)
		{
			let cc = CreateContentControl();
			let docContent = cc.GetContent();
			docContent.ClearContent(false);
			
			for (let iText = 0, nTexts = texts.length; iText < nTexts; ++iText)
			{
				let p = CreateParagraphWithText(texts[iText]);
				docContent.AddToContent(iText, p);
			}
			
			return cc;
		}
		logicDocument.AddToContent(0, AscTest.CreateParagraph());
		
		let cc1 = CreateFilledContentControl(["Text1", "Text2"]);
		let cc2 = CreateFilledContentControl(["Text3", "Text4"]);
		let p = CreateParagraphWithText("123");
		let lastPara = cc1.GetLastParagraph();
		let firstPara = cc2.GetFirstParagraph();
		
		logicDocument.AddToContent(0, cc1);
		logicDocument.AddToContent(1, p);
		logicDocument.AddToContent(2, cc2);
		
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");

		AscTest.MoveCursorToParagraph(p, true);
		AscTest.PressKey(AscTest.Key.backspace);
		assert.ok(true, "Move to the start of the middle paragraph and click backspace button");
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), true, "Check if paragraph is present in the document");
		assert.strictEqual(lastPara.IsThisElementCurrent() && lastPara.IsCursorAtEnd(), true, "Check cursor position at the end of the first content control");
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.PressKey(AscTest.Key.delete);
		assert.ok(true, "Move to the end of the middle paragraph and click delete button");
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), true, "Check if paragraph is present in the document");
		assert.strictEqual(firstPara.IsThisElementCurrent() && firstPara.IsCursorAtBegin(), true, "Check cursor position at the start of the second content control");
		
		AscTest.ClearParagraph(p);
		
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.PressKey(AscTest.Key.backspace);
		assert.ok(true, "Move to the start of the middle paragraph and click backspace button");
		assert.strictEqual(logicDocument.GetElementsCount(), 3, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), false, "Check if paragraph is present in the document");
		assert.strictEqual(lastPara.IsThisElementCurrent() && lastPara.IsCursorAtEnd(), true, "Check cursor position at the end of the first content control");
		
		logicDocument.AddToContent(1, p);
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.PressKey(AscTest.Key.delete);
		assert.ok(true, "Move to the end of the middle paragraph and click delete button");
		assert.strictEqual(logicDocument.GetElementsCount(), 3, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), false, "Check if paragraph is present in the document");
		assert.strictEqual(firstPara.IsThisElementCurrent() && firstPara.IsCursorAtBegin(), true, "Check cursor position at the start of the second content control");
		
		logicDocument.AddToContent(1, p);

		AscTest.SetTrackRevisions(true);
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.PressKey(AscTest.Key.backspace);
		
		assert.strictEqual(logicDocument.IsTrackRevisions(), true, "Turn on track revisions");
		assert.ok(true, "Move to the start of the middle paragraph and click backspace button");
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), true, "Check if paragraph is present in the document");
		assert.strictEqual(lastPara.IsThisElementCurrent() && lastPara.IsCursorAtEnd(), true, "Check cursor position at the end of the first content control");
		assert.strictEqual(lastPara.GetReviewType(), reviewtype_Remove, "Check that the last paragraph in first cc has become deleted on review");
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.PressKey(AscTest.Key.delete);
		assert.ok(true, "Move to the end of the middle paragraph and click delete button");
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), true, "Check if paragraph is present in the document");
		assert.strictEqual(firstPara.IsThisElementCurrent() && firstPara.IsCursorAtBegin(), true, "Check cursor position at the start of the second content control");
		assert.strictEqual(p.GetReviewType(), reviewtype_Remove, "Check that middle paragraph has become deleted on review");
		
		p.SetReviewType(reviewtype_Add);
		assert.strictEqual(p.GetReviewType(), reviewtype_Add, "Change review type of the middle paragraph to added on review");
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.PressKey(AscTest.Key.delete);
		assert.ok(true, "Move to the end of the middle paragraph and click delete button");
		assert.strictEqual(logicDocument.GetElementsCount(), 3, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), false, "Check if paragraph is present in the document");
		assert.strictEqual(firstPara.IsThisElementCurrent() && firstPara.IsCursorAtBegin(), true, "Check cursor position at the start of the second content control");
	});
	
});
