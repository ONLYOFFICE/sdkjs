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
	const logicDocument    = AscTest.CreateLogicDocument();
	const styleManager     = logicDocument.GetStyleManager();
	const numberingManager = logicDocument.GetNumberingManager();
	
	QUnit.module("Test the application of numbering to the document");
	
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
});
