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

});
