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
	const pageH  = 1000;
	const pageW  = 500;
	const marginT = 50;
	const marginB = 50;
	const marginR = 50;
	const marginL = 50;
	
	const logicDocument = AscTest.CreateLogicDocument();
	
	function SetupDocumentSection()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(pageW, pageH);
		sectPr.SetPageMargins(marginL, marginT, marginR, marginB);
	}
	SetupDocumentSection();
	
	function CreateImageInParagraph(p, w, h)
	{
		let d = AscTest.CreateImage(w, h);
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		run.AddToContent(0, d);
		return d;
	}
	
	function AddTextToParagraph(p, text)
	{
		let run = new AscWord.CRun();
		p.AddToContentToEnd(run);
		run.AddText(text);
	}
	
	QUnit.module("Test the positioning of floating drawings");
	
	QUnit.test("Test bugs #50253 #61936", function (assert)
	{
		let p1 = AscTest.CreateParagraph();
		let p2 = AscTest.CreateParagraph();
		
		p1.SetParagraphSpacing({Before : 0, After : 0, LineRule : linerule_Auto, Line : 1});
		
		let d1 = CreateImageInParagraph(p1, 50, 50);
		let d2 = CreateImageInParagraph(p2, 50, 50);
		
		AscTest.ClearDocument();
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		
		assert.strictEqual(logicDocument.GetElementsCount(), 2, "Should be 2 paragraphs in the document");
		
		d1.Set_DrawingType(drawing_Anchor);
		d1.Set_Distance(10, 10, 10, 10);
		d1.Set_WrappingType(WRAPPING_TYPE_SQUARE);
		d1.Set_PositionH(Asc.c_oAscRelativeFromH.Page, false, 40);
		d1.Set_PositionV(Asc.c_oAscRelativeFromV.Page, false, 40);
		
		d2.Set_DrawingType(drawing_Anchor);
		d2.Set_PositionH(Asc.c_oAscRelativeFromH.Column, false, 0);
		d2.Set_PositionV(Asc.c_oAscRelativeFromV.Paragraph, false, 0);
		d2.Set_WrappingType(WRAPPING_TYPE_NONE);
		d2.Set_BehindDoc(false);
		
		function CheckPositions()
		{
			AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word14);
			AscTest.Recalculate();
			assert.ok(true, "Set compatibility mode equal to 14 (Word2010)");
			assert.strictEqual(d1.X, 40, "Check drawing1.x");
			assert.strictEqual(d1.Y, 40, "Check drawing1.y");
			assert.strictEqual(d2.X, 100, "Check drawing2.x");
			assert.strictEqual(d2.Y, 50 + AscTest.FontHeight, "Check drawing2.y");
			
			AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
			AscTest.Recalculate();
			assert.ok(true, "Set compatibility mode equal to 15 (Word2013-2019)");
			assert.strictEqual(d1.X, 40, "Check drawing1.x");
			assert.strictEqual(d1.Y, 40, "Check drawing1.y");
			assert.strictEqual(d2.X, 50, "Check drawing2.x");
			assert.strictEqual(d2.Y, 50 + AscTest.FontHeight, "Check drawing2.y");
		}
		
		assert.ok(true, "Both paragraphs have no text");
		CheckPositions();
		
		AddTextToParagraph(p1, "First");
		AddTextToParagraph(p2, "Second");
		
		assert.ok(true, "Fill in paragraphs with text");
		CheckPositions();

		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
	});
	
	QUnit.test("Check for a case when a text-wrapped object  extends beyond the boundaries of the fields (#68037)", function (assert)
	{
		let p1 = AscTest.CreateParagraph();
		let p2 = AscTest.CreateParagraph();
		
		p1.SetParagraphSpacing({Before : 0, After : pageH - AscTest.FontHeight - 70 - marginT - marginB, LineRule : linerule_Auto, Line : 1});
		
		let p1Bottom = pageH - 70 - marginB;
		
		let d1 = CreateImageInParagraph(p2, 50, 50);
		
		AscTest.ClearDocument();
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		
		assert.strictEqual(logicDocument.GetElementsCount(), 2, "Should be 2 paragraphs in the document");
		
		let offsetY = 40;
		
		d1.Set_DrawingType(drawing_Anchor);
		d1.Set_PositionH(Asc.c_oAscRelativeFromH.Column, false, 0);
		d1.Set_PositionV(Asc.c_oAscRelativeFromV.Paragraph, false, offsetY);
		d1.Set_WrappingType(WRAPPING_TYPE_SQUARE);
		d1.Set_BehindDoc(false);
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word14);
		AscTest.Recalculate();
		assert.ok(true, "Set compatibility mode equal to 14 (Word2010)");
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Check pages count");
		assert.strictEqual(d1.X, marginL, "Check drawing1.x");
		assert.strictEqual(d1.Y, p1Bottom + offsetY, "Check drawing1.y");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		AscTest.Recalculate();
		assert.ok(true, "Set compatibility mode equal to 15 (Word2013-2019)");
		assert.strictEqual(logicDocument.GetPagesCount(), 2, "Check pages count");
		assert.strictEqual(d1.X, marginL, "Check drawing1.x");
		assert.strictEqual(d1.Y, marginT + offsetY, "Check drawing1.y");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
	});
	
});
