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
	const logicDocument    = AscTest.CreateLogicDocument();
	const styleManager     = logicDocument.GetStyleManager();
	const numberingManager = logicDocument.GetNumberingManager();
	
	QUnit.module("Test the application of numbering to the document");
	
	let styleCounter = 0;
	function CreateStyle()
	{
		let style = new AscWord.CStyle("style" + (++styleCounter), null, null, styletype_Paragraph);
		styleManager.Add(style);
		return style;
	}
	
	function CreateNum()
	{
		let numInfo = AscWord.GetNumberingObjectByDeprecatedTypes(2, 7);
		let num = numberingManager.CreateNum();
		numInfo.FillNum(num);
		numberingManager.AddNum(num);
		return num;
	}
	
	QUnit.test("Apply numbering through style", function (assert)
	{
		function AddParagraph(style, text)
		{
			let p = AscTest.CreateParagraph();
			logicDocument.PushToContent(p);
			p.SetParagraphStyle(style.GetName());
			
			let run = new AscWord.CRun();
			p.AddToContent(0, run);
			run.AddText(text);
			return p;
		}
		
		function CheckParagraph(paraIndex, text)
		{
			let p = logicDocument.GetElement(paraIndex);
			assert.strictEqual(p.GetNumberingText(false), text, "Check numbering text for paragraph " + paraIndex);
		}
		
		let style = CreateStyle();
		
		AscTest.ClearDocument();
		let p1 = AddParagraph(style, "First");
		let p2 = AddParagraph(style, "Second");
		let p3 = AddParagraph(style, "Third");
		
		//--------------------------------------------------------------------------------------------------------------
		// Нет нумерации
		//--------------------------------------------------------------------------------------------------------------
		AscTest.Recalculate();
		CheckParagraph(0, "");
		CheckParagraph(1, "");
		CheckParagraph(2, "");
		
		//--------------------------------------------------------------------------------------------------------------
		// Нумерация указана в стиле, и в стиле сразу заданы правильные уровни
		//--------------------------------------------------------------------------------------------------------------
		let num = CreateNum();
		num.GetLvl(0).SetPStyle(style.GetId());
		style.SetNumPr(num.GetId(), 0);
		AscTest.Recalculate();
		
		CheckParagraph(0, "1.");
		CheckParagraph(1, "2.");
		CheckParagraph(2, "3.");
		//--------------------------------------------------------------------------------------------------------------
		// Перемещаем курсор во второй параграф и меняем нумерацию на a) b) c)
		//--------------------------------------------------------------------------------------------------------------
		AscTest.MoveCursorToParagraph(p2, true);
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(1, 5));
		AscTest.Recalculate();
		
		CheckParagraph(0, "a)");
		CheckParagraph(1, "b)");
		CheckParagraph(2, "c)");
	});
	
	QUnit.test("Numbering for headings", function (assert)
	{
		function CheckHeading(iLvl, text)
		{
			let p     = logicDocument.GetElement(iLvl);
			let numPr = p.GetNumPr();
			let num   = numberingManager.GetNum(numPr.NumId);
			if (!num)
				return assert.strictEqual(false, true, "No numbering in heading " + (iLvl + 1));

			let numLvl  = num.GetLvl(iLvl);
			let styleId = styleManager.GetDefaultHeading(iLvl);
			let style   = styleManager.Get(styleId);
			if (!style)
				return assert.strictEqual(false, true, "No style for heading " + (iLvl + 1));
			
			if (!style.ParaPr.NumPr)
				return assert.strictEqual(false, true, "No numbering in style for heading " + (iLvl + 1));
			
			assert.strictEqual(p.GetNumberingText(false), text, "Check numbering text for heading " + (iLvl + 1));
			assert.strictEqual(numLvl.GetPStyle(), styleId, "Check heading style in numbering for heading " + (iLvl + 1));
			assert.deepEqual(numPr, style.ParaPr.NumPr, "Check numbering in heading style for heading " + (iLvl + 1));
		}
		
		function CheckNoNumbering(iLvl)
		{
			let p = logicDocument.GetElement(iLvl);
			assert.strictEqual(p.HaveNumbering(), false, "Check numbering in heading " + (iLvl + 1) + " after its removal");
			let headingStyleId = styleManager.GetDefaultHeading(iLvl);
			let pStyleId = p.GetParagraphStyle();
			assert.strictEqual(headingStyleId, pStyleId, "Check style remaining style in paragraph " + (iLvl + 1));
			let style = styleManager.Get(pStyleId);
			assert.strictEqual(style.HaveNumbering(), true, "Check numbering in heading" + (iLvl + 1) + " style");
		}
		
		AscTest.ClearDocument();
		
		for (let iHead = 0; iHead < 9; ++iHead)
		{
			let p = AscTest.CreateParagraph();
			logicDocument.PushToContent(p);
			let styleId = styleManager.GetDefaultHeading(iHead);
			p.SetParagraphStyle(styleManager.Get(styleId).GetName());
			let run = new AscWord.CRun();
			p.AddToContent(0, run);
			run.AddText("Heading " + iHead);
		}
		
		assert.strictEqual(logicDocument.GetElementsCount(), 9, "Check number of paragraphs");
		logicDocument.SelectAll();
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(2, 4));
		AscTest.Recalculate();
		
		CheckHeading(0, "Article I.");
		CheckHeading(1, "Section I.01");
		CheckHeading(2, "(a)");
		CheckHeading(3, "(i)");
		CheckHeading(4, "1)");
		CheckHeading(5, "a)");
		CheckHeading(6, "i)");
		CheckHeading(7, "a.");
		CheckHeading(8, "i.");
		
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(2, 7));
		AscTest.Recalculate();
		
		CheckHeading(0, "1.");
		CheckHeading(1, "1.1.");
		CheckHeading(2, "1.1.1.");
		CheckHeading(3, "1.1.1.1.");
		CheckHeading(4, "1.1.1.1.1.");
		CheckHeading(5, "1.1.1.1.1.1.");
		CheckHeading(6, "1.1.1.1.1.1.1.");
		CheckHeading(7, "1.1.1.1.1.1.1.1.");
		CheckHeading(8, "1.1.1.1.1.1.1.1.1.");
		
		// Cancel numbering (heading still have numbering in it's style)
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(2, -1));
		AscTest.Recalculate();
		
		for (let iLvl = 0; iLvl < 9; ++iLvl)
			CheckNoNumbering(iLvl);
		
		// Re-apply numbering to list with canceled numbering
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(2, 7));
		AscTest.Recalculate();
		
		CheckHeading(0, "1.");
		CheckHeading(1, "1.1.");
		CheckHeading(2, "1.1.1.");
		CheckHeading(3, "1.1.1.1.");
		CheckHeading(4, "1.1.1.1.1.");
		CheckHeading(5, "1.1.1.1.1.1.");
		CheckHeading(6, "1.1.1.1.1.1.1.");
		CheckHeading(7, "1.1.1.1.1.1.1.1.");
		CheckHeading(8, "1.1.1.1.1.1.1.1.1.");
	});
	
	QUnit.test("Applying numbering by selecting vs. placing cursor", function (assert)
	{
		function CheckParagraph(paraIndex, text)
		{
			let p = logicDocument.GetElement(paraIndex);
			assert.strictEqual(p.GetNumberingText(false), text, "Check numbering text for paragraph " + paraIndex);
		}
		
		function GenerateDocument()
		{
			AscTest.ClearDocument();
			
			for (let i = 0; i < 2; ++i)
			{
				for (let iHead = 0; iHead < 4; ++iHead)
				{
					let p = AscTest.CreateParagraph();
					logicDocument.PushToContent(p);
					let styleId = styleManager.GetDefaultHeading(iHead);
					p.SetParagraphStyle(styleManager.Get(styleId).GetName());
					let run = new AscWord.CRun();
					p.AddToContent(0, run);
					run.AddText("Heading " + iHead);
				}
			}
			assert.strictEqual(logicDocument.GetElementsCount(), 8, "Check number of paragraphs");
		}
		
		GenerateDocument();
		logicDocument.SelectAll();
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(2, 4));
		AscTest.Recalculate();
		
		CheckParagraph(0, "Article I.");
		CheckParagraph(1, "Section I.01");
		CheckParagraph(2, "(a)");
		CheckParagraph(3, "(i)");
		CheckParagraph(4, "Article II.");
		CheckParagraph(5, "Section II.01");
		CheckParagraph(6, "(a)");
		CheckParagraph(7, "(i)");
		
		
		// Check applying numbering when cursor is placed in one of paragraphs
		logicDocument.RemoveSelection();
		let p = logicDocument.GetElement(1);
		AscTest.MoveCursorToParagraph(p, true);
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(2, 7));
		AscTest.Recalculate();
		
		CheckParagraph(0, "1.");
		CheckParagraph(1, "1.1.");
		CheckParagraph(2, "1.1.1.");
		CheckParagraph(3, "1.1.1.1.");
		CheckParagraph(4, "2.");
		CheckParagraph(5, "2.1.");
		CheckParagraph(6, "2.1.1.");
		CheckParagraph(7, "2.1.1.1.");
		
		// Применяем нумерацию с Heading по селекту (должно примениться ко всем параграфам, даже не выделенным)
		AscTest.SelectDocumentRange(0, 3);
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(2, 6));
		
		CheckParagraph(0, "I.");
		CheckParagraph(1, "A.");
		CheckParagraph(2, "1.");
		CheckParagraph(3, "a)");
		CheckParagraph(4, "II.");
		CheckParagraph(5, "A.");
		CheckParagraph(6, "1.");
		CheckParagraph(7, "a)");
		
		// Применяем нумерацию буз Heading по селекту (должно примениться только к выделенным параграфам)
		AscTest.SelectDocumentRange(0, 3);
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(2, 2));
		
		CheckParagraph(0, "1.");
		CheckParagraph(1, "1.1.");
		CheckParagraph(2, "1.1.1.");
		CheckParagraph(3, "1.1.1.1.");
		CheckParagraph(4, "I.");
		CheckParagraph(5, "A.");
		CheckParagraph(6, "1.");
		CheckParagraph(7, "a)");
	});
	
	QUnit.test("Applying numbering by selecting to paragraphs with left indentation", function (assert)
	{
		// Check situation from bug 71056
		
		function CheckParagraph(paraIndex, text)
		{
			let p = logicDocument.GetElement(paraIndex);
			assert.strictEqual(p.GetNumberingText(false), text, "Check numbering text for paragraph " + paraIndex);
		}
		
		function GenerateDocument()
		{
			AscTest.ClearDocument();
			
			for (let i = 0; i < 4; ++i)
			{
				let p = AscTest.CreateParagraph();
				p.SetParagraphIndent({Left : 5, FirstLine : 15});
				logicDocument.PushToContent(p);
				let run = new AscWord.CRun();
				p.AddToContent(0, run);
				run.AddText("Paragraph " + i);
			}
			assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of paragraphs");
		}
		
		GenerateDocument();
		logicDocument.SelectAll();
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(1, 2));
		AscTest.Recalculate();
		
		CheckParagraph(0, "1)");
		CheckParagraph(1, "2)");
		CheckParagraph(2, "3)");
		CheckParagraph(3, "4)");
		
		GenerateDocument();
		logicDocument.GetElement(0).SetParagraphIndent({Left : 5, FirstLine : 25});
		logicDocument.SelectAll();
		logicDocument.SetParagraphNumbering(AscWord.GetNumberingObjectByDeprecatedTypes(1, 2));
		AscTest.Recalculate();
		
		CheckParagraph(0, "i.");
		CheckParagraph(1, "b.");
		CheckParagraph(2, "c.");
		CheckParagraph(3, "d.");
	});
});
