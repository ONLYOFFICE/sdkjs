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

	const charWidth = AscTest.CharWidth * AscTest.FontSize;

	let dc = new AscWord.CDocumentContent();
	dc.ClearContent(false);

	let para = new AscWord.Paragraph();
	dc.AddToContent(0, para);

	let run = new AscWord.CRun();
	para.AddToContent(0, run);

	function recalculate(width)
	{
		dc.Reset(0, 0, width, 10000);
		dc.Recalculate_Page(0, true);
	}

	function setText(text)
	{
		run.ClearContent();
		run.AddText(text);
	}
	
	function checkLines(assert, para, lines)
	{
		assert.strictEqual(para.GetLinesCount(), lines.length, "Check number of lines");
		
		for (let i = 0, lineCount = Math.min(lines.length, para.GetLinesCount()); i < lineCount; ++i)
		{
			assert.strictEqual(para.GetTextOnLine(i), lines[i], "Check text on line " + i);
		}
	}

	QUnit.module("Paragraph Lines");

	QUnit.test("Test regular line break cases", function (assert)
	{
		setText("1234");
		recalculate(charWidth * 3.5);
		checkLines(assert, para, [
			"123",
			"4"
		]);
		
		setText("12 34");
		recalculate(charWidth * 3.5);
		checkLines(assert, para, [
			"12 ",
			"34"
		]);
	});
	
	QUnit.test("Test line breaks for Asian text", function (assert)
	{
		setText("你好世界! 你好世界! 你好世界! 你好世界! ");
		recalculate(charWidth * 8.5);
		checkLines(assert, para, [
			"你好世界! 你好",
			"世界! 你好世",
			"界! 你好世界! "
		]);
		
		setText("你好世界! Hello! 你好世界! 你好世界! ");
		recalculate(charWidth * 8.5);
		checkLines(assert, para, [
			"你好世界! ",
			"Hello! 你",
			"好世界! 你好世",
			"界! "
		]);
		
		// check non asian text with the eastAsian hint (71108)
		run.SetRFontsHint(AscWord.fonthint_EastAsia);
		setText("你好世界! Hello! 你好世界! 你好世界! ");
		recalculate(charWidth * 8.5);
		checkLines(assert, para, [
			"你好世界! He",
			"llo! 你好世",
			"界! 你好世界! "
		]);
		
		run.SetRFontsHint(undefined);
	});
	

	QUnit.test("Test: \"Test paragraph with very narrow width\"", function (assert)
	{
		assert.strictEqual(dc.GetElementsCount(), 1, "Paragraphs count");

		let narrowWidth = charWidth / 2;
		
		setText("");
		recalculate(narrowWidth);
		assert.strictEqual(para.GetLinesCount(), 1, "Lines count of empty paragraph");
		assert.deepEqual(para.GetLineBounds(0), new AscWord.CDocumentBounds(0, 0, narrowWidth, AscTest.FontHeight), "Line bounds of empty paragraph");
		assert.strictEqual(para.GetPagesCount(), 1, "Pages count of paragraph");
		
		setText("123");
		recalculate(narrowWidth);
		assert.strictEqual(AscTest.GetParagraphText(para), "123", "Paragraph text: 123");
		assert.strictEqual(para.GetLinesCount(), 3, "Lines count 3");
		assert.deepEqual(para.GetLineBounds(0), new AscWord.CDocumentBounds(0, 0, narrowWidth, AscTest.FontHeight), "Check line bounds");
		assert.deepEqual(para.GetLineBounds(1), new AscWord.CDocumentBounds(0, AscTest.FontHeight, narrowWidth, AscTest.FontHeight * 2), "Check line bounds");
		assert.deepEqual(para.GetLineBounds(2), new AscWord.CDocumentBounds(0, AscTest.FontHeight * 2, narrowWidth, AscTest.FontHeight * 3), "Check line bounds");
		assert.strictEqual(para.GetPagesCount(), 1, "Pages count of paragraph");
		
		setText("Q");
		recalculate(narrowWidth);
		assert.strictEqual(AscTest.GetParagraphText(para), "Q", "Paragraph text: Q");
		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.deepEqual(para.GetLineBounds(0), new AscWord.CDocumentBounds(0, 0, narrowWidth, AscTest.FontHeight), "Check line bounds");
		assert.strictEqual(para.GetPagesCount(), 1, "Pages count of paragraph");
		
		setText("Q ");
		recalculate(narrowWidth);
		assert.strictEqual(AscTest.GetParagraphText(para), "Q ", "Paragraph text: Q'<'space'>'");
		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.deepEqual(para.GetLineBounds(0), new AscWord.CDocumentBounds(0, 0, narrowWidth, AscTest.FontHeight), "Check line bounds");
		assert.strictEqual(para.GetPagesCount(), 1, "Pages count of paragraph");
	});

	QUnit.test("Test: \"Test line break of ligatures\"", function (assert)
	{
		setText("ffi");

		let text_f_0 = run.GetElement(0);
		let text_f_1 = run.GetElement(1);
		let text_i   = run.GetElement(2);

		recalculate(charWidth * 3.5);
		assert.strictEqual(AscTest.GetParagraphText(para), "ffi", "Paragraph text: ffi");

		assert.strictEqual(text_f_0.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE, "Check f code point type");
		assert.strictEqual(text_f_1.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE, "Check f code point type");
		assert.strictEqual(text_i.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE, "Check i code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "ffi", "Text on line 0 'ffi ");

		recalculate(charWidth * 2.5);

		assert.strictEqual(text_f_0.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE, "Check f code point type");
		assert.strictEqual(text_f_1.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE, "Check f code point type");
		assert.strictEqual(text_i.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check i code point type");

		assert.strictEqual(para.GetLinesCount(), 2, "Lines count 2");
		assert.strictEqual(para.GetTextOnLine(0), "ff", "Text on line 0 'ff'");
		assert.strictEqual(para.GetTextOnLine(1), "i", "Text on line 1 'i'");

		recalculate(charWidth * 1.5);

		assert.strictEqual(text_f_0.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check f code point type");
		assert.strictEqual(text_f_1.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check f code point type");
		assert.strictEqual(text_i.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check i code point type");

		assert.strictEqual(para.GetLinesCount(), 3, "Lines count 3");
		assert.strictEqual(para.GetTextOnLine(0), "f", "Text on line 0 'f'");
		assert.strictEqual(para.GetTextOnLine(1), "f", "Text on line 1 'f'");
		assert.strictEqual(para.GetTextOnLine(2), "i", "Text on line 2 'i'");
	});

	QUnit.test("Test: \"Test line break of combining marks\"", function (assert)
	{
		setText("xyz");

		let text_x = run.GetElement(0);
		let text_y = run.GetElement(1);
		let text_z = run.GetElement(2);

		recalculate(charWidth * 3.5);
		assert.strictEqual(AscTest.GetParagraphText(para), "xyz", "Paragraph text: xyz");

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz ");

		// Комбинированные символы НЕ ДОЛЖНЫ отдельно переносится на новую строку
		recalculate(charWidth * 2.5);

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz'");

		recalculate(charWidth * 1.5);

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz'");

		recalculate(charWidth * 1.5);

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz'");

		recalculate(charWidth * 0.5);

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz'");
	});
	
	QUnit.test("Test edge cases", function (assert)
	{
		// -------------------------------------------------------------------------------------------------------------
		// Check bug (76557)
		// We can make a break after '-'
		// But '%',')' can't be at the start of the line
		// so we should treat the whole text as a long word
		// '(' can be at the start of the line
		setText("hhhh-%hhh");
		recalculate(charWidth * 6.5);
		checkLines(assert, para, [
			"hhhh-%",
			"hhh"
		]);
		
		setText("hhhh-)hhh");
		recalculate(charWidth * 6.5);
		checkLines(assert, para, [
			"hhhh-)",
			"hhh"
		]);
		
		setText("hhhh-(hhh");
		recalculate(charWidth * 6.5);
		checkLines(assert, para, [
			"hhhh-",
			"(hhh"
		]);
		// -------------------------------------------------------------------------------------------------------------
	});

});
