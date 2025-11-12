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


$(function () {
	
	let logicDocument    = AscTest.CreateLogicDocument();
	let styleManager     = logicDocument.GetStyles();
	let numberingManager = logicDocument.GetNumbering();
	
	function AddParagraph(style, text)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		p.SetParagraphStyleById(style.GetId());
		
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		run.AddText(text);
		return p;
	}
	
	let styleCounter = 0;
	function CreateStyle()
	{
		let style = new AscWord.CStyle("style" + (++styleCounter), null, null, styletype_Paragraph);
		styleManager.Add(style);
		return style;
	}
	
	function CreateNum()
	{
		let num = numberingManager.CreateNum();
		numberingManager.AddNum(num);
		let numLvl = num.GetLvl(0).Copy();

		let paraPr = numLvl.GetParaPr();
		paraPr.Ind.Left      = 15;
		paraPr.Ind.FirstLine = 5;
		numLvl.SetParaPr(paraPr);
		
		num.SetLvl(numLvl, 0);
		
		return num;
	}
	
	QUnit.module("Style applicator");
	
	QUnit.test("Style application and change from current selection", function (assert)
	{
		AscTest.ClearDocument();
		
		let style = CreateStyle();
		p1 = AddParagraph(style, "First");
		p2 = AddParagraph(style, "Second");
		
		let num = CreateNum();
		
		function TestParagraph(p, left, right, first, align)
		{
			let compiledPr = p.GetCompiledParaPr(false);
			assert.strictEqual(compiledPr.Ind.Left, left, "Check Left indent");
			assert.strictEqual(compiledPr.Ind.Right, right, "Check Right indent");
			assert.strictEqual(compiledPr.Ind.FirstLine, first, "Check FirstLine");
			assert.strictEqual(compiledPr.Jc, align, "Check align");
		}
		
		assert.ok(true, "Create empty style and set it to two paragraphs. Check default values");
		
		TestParagraph(p1, 0, 0, 0, AscCommon.align_Left);
		TestParagraph(p2, 0, 0, 0, AscCommon.align_Left);
		
		
		assert.ok(true, "Apply direct properties and check values");
		p1.SetParagraphIndent({Left : 10, Right : 10, FirstLine : 20});
		p1.SetParagraphAlign(AscCommon.align_Center);
		
		TestParagraph(p1, 10, 10, 20, AscCommon.align_Center);

		assert.ok(true, "Apply direct numPr and check indents");
		p1.SetNumPr(num.GetId(), 0);
		p2.SetNumPr(num.GetId(), 0);
		
		TestParagraph(p1, 10, 10, 20, AscCommon.align_Center);
		TestParagraph(p2, 15, 0, 5, AscCommon.align_Left);
		
		// Обновление стиля в интерфейсе сейчас работает, как получение текущего по выделению, дальше ему проставляется точно такое же имя
		// и мы его заново выставляем в документ
		let uiStyle = logicDocument.GetStyleFromFormatting();
		uiStyle.put_Name(style.GetName());
		AscTest.MoveCursorToParagraph(p1, true);
		logicDocument.Add_NewStyle(uiStyle);
		
		assert.ok(true, "Update style by paraPr of the first paragraph and check paragraph compiled properties after that");
		TestParagraph(p1, 10, 10, 20, AscCommon.align_Center);
		TestParagraph(p2, 10, 10, 20, AscCommon.align_Center);
	});
});
