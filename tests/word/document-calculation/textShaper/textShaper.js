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

$(function () {

	AscTest.CreateLogicDocument();

	const charWidth = AscTest.CharWidth * AscTest.FontSize;

	let dc = new AscWord.CDocumentContent();
	dc.ClearContent(false);

	let para = AscTest.CreateParagraph();
	dc.AddToContent(0, para);

	let run = new AscWord.CRun();
	para.AddToContent(0, run);

	function Recalculate(width)
	{
		if (!width)
			width = 100 * charWidth;

		dc.Reset(0, 0, width, 10000);
		dc.Recalculate_Page(0, true);
	}

	function SetText(text)
	{
		run.ClearContent();
		run.AddText(text);
	}

	function GetText()
	{
		return AscTest.GetParagraphText(para);
	}

	function TestText(assert, text)
	{
		SetText(text);
		Recalculate();

		assert.strictEqual(GetText(), text, "Paragraph text: " + text);
	}

	function TestCodePointType(assert, types)
	{
		let count = run.GetElementsCount();

		assert.strictEqual(count, types.length, "Check run element count");

		if (count !== types.length)
		{
			assert.true(false, "Bad elements and types length");
			return;
		}

		for (let index = 0; index < count; ++index)
		{
			let item = run.GetElement(index);
			assert.strictEqual(item.GetCodePointType(), types[index], "Check " + String.fromCodePoint(item.GetCodePoint()) + " code point type");
		}
	}

	function TestCursorMove(assert, count)
	{
		para.MoveCursorToStartPos();

		for (let index = 0; index < count; ++index)
		{
			assert.strictEqual(para.IsCursorAtEnd(), false, "Check cursor move right " + index);
			para.MoveCursorRight();
		}

		assert.strictEqual(para.IsCursorAtEnd(), true, "Check cursor at the end");

		para.MoveCursorToEndPos();
		assert.strictEqual(para.IsCursorAtEnd(), true, "Check cursor at the end");

		for (let index = 0; index < count; ++index)
		{
			assert.strictEqual(para.IsCursorAtBegin(), false, "Check cursor move left " + index);
			para.MoveCursorLeft();
		}

		assert.strictEqual(para.IsCursorAtBegin(), true, "Check cursor at the start");
	}

	function TestDelete(assert, text, countRemove, countDelete)
	{
		TestText(assert, text);
		para.MoveCursorToStartPos();

		for (let index = 0; index < countDelete; ++index)
		{
			assert.strictEqual(para.IsEmpty(), false, "Check delete " + index);
			para.Remove(1);
		}

		assert.strictEqual(para.IsEmpty(), true, "Check end of delete");

		TestText(assert, text);
		para.MoveCursorToEndPos();

		for (let index = 0; index < countRemove; ++index)
		{
			assert.strictEqual(para.IsEmpty(), false, "Check remove " + index);
			para.Remove(-1);
		}

		assert.strictEqual(para.IsEmpty(), true, "Check end of remove");

	}

	QUnit.module("Text shaper");

	QUnit.test("Test: \"code point types\"", function (assert)
	{
		function Test(text, codePointTypes, moveCount, removeCount, deleteCount)
		{
			TestText(assert, text);
			TestCodePointType(assert, codePointTypes);
			TestCursorMove(assert, moveCount);
			TestDelete(assert, text, removeCount, deleteCount);
		}

		Test("abc",
			[AscWord.CODEPOINT_TYPE.BASE, AscWord.CODEPOINT_TYPE.BASE, AscWord.CODEPOINT_TYPE.BASE],
			3,
			3,
			3);

		Test("ffi",
			[AscWord.CODEPOINT_TYPE.LIGATURE, AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE, AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE],
			3,
			3,
			3);

		Test("xyz",
			[AscWord.CODEPOINT_TYPE.BASE, AscWord.CODEPOINT_TYPE.COMBINING_MARK, AscWord.CODEPOINT_TYPE.COMBINING_MARK],
			1,
			3,
			1);

		// Проверяем диакритический знак, который не складывается шейпером
		Test("á",
			[AscWord.CODEPOINT_TYPE.BASE, AscWord.CODEPOINT_TYPE.BASE],
			1,
			2,
			1);
	});
});
