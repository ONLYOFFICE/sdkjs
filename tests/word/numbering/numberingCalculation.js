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
	
	
	QUnit.module("Test the numbering calculation");
	
	QUnit.test("Test the numbering specified in a style", function (assert)
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
		
		// Задаем 3 независимых стиля и свяжем их с тремя уровнями нумерации
		let style0 = CreateStyle();
		let style1 = CreateStyle();
		let style2 = CreateStyle();
		
		function GenerateDocument()
		{
			AscTest.ClearDocument();
			AddParagraph(style0, "Style1");
			AddParagraph(style1, "Style2");
			AddParagraph(style2, "Style3");
		}
		
		function Recalculate()
		{
			GenerateDocument();
			AscTest.Recalculate();
		}
		
		//--------------------------------------------------------------------------------------------------------------
		// Нет нумерации
		//--------------------------------------------------------------------------------------------------------------
		Recalculate();
		CheckParagraph(0, "");
		CheckParagraph(1, "");
		CheckParagraph(2, "");
		
		//--------------------------------------------------------------------------------------------------------------
		// Нумерация указана в стиле, и в стиле сразу заданы правильные уровни
		//--------------------------------------------------------------------------------------------------------------
		let num = CreateNum();
		num.GetLvl(0).SetPStyle(style0.GetId());
		num.GetLvl(1).SetPStyle(style1.GetId());
		num.GetLvl(2).SetPStyle(style2.GetId());
		
		style0.SetNumPr(num.GetId(), 0);
		style1.SetNumPr(num.GetId(), 1);
		style2.SetNumPr(num.GetId(), 2);
		Recalculate();
		
		CheckParagraph(0, "1.");
		CheckParagraph(1, "1.1.");
		CheckParagraph(2, "1.1.1.");
		//--------------------------------------------------------------------------------------------------------------
		// Нумерация указана в стиле, но в стиле уровни не указаны
		// Не смотря на то, что в спецификации написано, что мы должны определять уровень по pStyle в уровне нумерации,
		// MSWord так не делает. В MSWord если не задан уровнь, значит уровень = 0. А если стиль не совпадает со стилем
		// в заданном уровне, значит нумерации нет.
		//--------------------------------------------------------------------------------------------------------------
		style0.SetNumPr(num.GetId(), undefined);
		style1.SetNumPr(num.GetId(), undefined);
		style2.SetNumPr(num.GetId(), undefined);
		
		Recalculate();
		CheckParagraph(0, "1.");
		CheckParagraph(1, "");
		CheckParagraph(2, "");
		//--------------------------------------------------------------------------------------------------------------
		// Нумерация указана в стиле, но в стиле уровни не указаны. Сами стили наследуются друг от друга
		// В добавок к предыдущей ситуации, MSWord проверят иерархию наследования стилей, если один из цепочки подходит
		// с текущим уровнем, то нумерация добавляется именно для подошедшего стиля.
		// https://bugzilla.onlyoffice.com/show_bug.cgi?id=51893
		//--------------------------------------------------------------------------------------------------------------
		style1.SetNumPr(null);
		style2.SetNumPr(null);
		style1.SetBasedOn(style0.GetId());
		style2.SetBasedOn(style1.GetId());
		
		Recalculate();
		CheckParagraph(0, "1.");
		CheckParagraph(1, "2.");
		CheckParagraph(2, "3.");
	});
	
	// TODO: Добавить больше тестов и вынести в отдельный файл
	QUnit.test("Test numbering collection", function (assert)
	{
		function AddParagraph(text)
		{
			let p = AscTest.CreateParagraph();
			logicDocument.PushToContent(p);
			
			let run = new AscWord.CRun();
			p.AddToContent(0, run);
			run.AddText(text);
			return p;
		}
		
		
		AscTest.ClearDocument();
		let p1 = AddParagraph("Paragraph 1");
		let p2 = AddParagraph("Paragraph 2");
		let p3 = AddParagraph("Paragraph 3");
		let p4 = AddParagraph("Paragraph 4");

		let style = CreateStyle();
		p2.SetParagraphStyle(style.GetName());
		
		let num = CreateNum();
		let numPr = new AscWord.NumPr(num.GetId(), 0);
		let collection = logicDocument.GetNumberingCollection();
		let paraArray = collection.GetAllParagraphsByNumbering(numPr);
		
		assert.strictEqual(paraArray.length, 0, "Check paragraphs for just created numbering");
		
		p1.SetNumPr(numPr.NumId, numPr.Lvl);
		paraArray = collection.GetAllParagraphsByNumbering(numPr);
		assert.strictEqual(paraArray.length, 1, "Add numbering direct to paragraph and check numbering collection");
		assert.ok(-1 !== paraArray.indexOf(p1), "Check paragraph in collection");
		assert.strictEqual(p1.GetNumberingText(false), "1.", "Check numbering text for paragraph " + 0);

		num.GetLvl(0).SetPStyle(style.GetId());
		style.SetNumPr(num.GetId(), 0);
		paraArray = collection.GetAllParagraphsByNumbering(numPr);
		assert.strictEqual(paraArray.length, 2, "Add numbering to a style. We have a predefined paragraph with that style. Check numbering collection");
		assert.ok(-1 !== paraArray.indexOf(p2), "Check paragraph in collection");
		assert.strictEqual(p2.GetNumberingText(false), "2.", "Check numbering text for paragraph " + 1);
		
		p3.SetParagraphStyle(style.GetName());
		paraArray = collection.GetAllParagraphsByNumbering(numPr);
		assert.strictEqual(paraArray.length, 3, "Add style with numbering to paragraph and check numbering collection");
		assert.ok(-1 !== paraArray.indexOf(p3), "Check paragraph in collection");
		assert.strictEqual(p3.GetNumberingText(false), "3.", "Check numbering text for paragraph " + 2);
		
		p4.SetParagraphStyle(style.GetName());
		paraArray = collection.GetAllParagraphsByNumbering(numPr);
		assert.strictEqual(paraArray.length, 4, "Add style with numbering to paragraph and check numbering collection");
		assert.ok(-1 !== paraArray.indexOf(p4), "Check paragraph in collection");
		assert.strictEqual(p4.GetNumberingText(false), "4.", "Check numbering text for paragraph " + 3);
		
		p4.SetNumPr(0, 0);
		paraArray = collection.GetAllParagraphsByNumbering(numPr);
		assert.strictEqual(paraArray.length, 3, "Cancel numbering by direct paragraph properties and check numbering collection");
		assert.ok(-1 === paraArray.indexOf(p4), "Check paragraph is not in the collection");
		assert.strictEqual(p4.GetNumberingText(false), "", "Check that there is no numbering text for paragraph " + 3);
		
		p3.SetParagraphStyleById(styleManager.GetDefaultParagraph());
		paraArray = collection.GetAllParagraphsByNumbering(numPr);
		assert.strictEqual(paraArray.length, 2, "Cancel numbering by removing style from paragraph. Check numbering collection");
		assert.ok(-1 === paraArray.indexOf(p3), "Check paragraph is not in the collection");
		assert.strictEqual(p3.GetNumberingText(false), "", "Check that there is no numbering text for paragraph " + 2);
		
		num.GetLvl(0).SetPStyle(undefined);
		style.SetNumPr(undefined);
		paraArray = collection.GetAllParagraphsByNumbering(numPr);
		assert.strictEqual(paraArray.length, 1, "Remove numbering link from style and check numbering collection");
		assert.ok(-1 === paraArray.indexOf(p2), "Check that paragraph with this style is not in the collection");
		assert.strictEqual(p2.GetNumberingText(false), "", "Check that there is no numbering text for paragraph " + 1);
		
		logicDocument.RemoveFromContent(0, 1);
		assert.strictEqual(p1.IsUseInDocument(), false, "Remove first paragraph and check if it is not present in the document");
		paraArray = collection.GetAllParagraphsByNumbering(numPr);
		assert.strictEqual(paraArray.length, 0, "Remove first paragraph and check numbering collection");
		assert.ok(-1 === paraArray.indexOf(p1), "Check that removed paragraph is not in the collection");
	});

});
