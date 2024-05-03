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

	function Recalculate(width)
	{
		dc.Reset(0, 0, width, 10000);
		dc.Recalculate_Page(0, true);
	}

	function SetText(text)
	{
		run.ClearContent();
		run.AddText(text);
	}

	QUnit.module("Paragraph Lines");

	QUnit.test("Test: \"Test regular line break cases\"", function (assert)
	{
		SetText("1234");
		Recalculate(charWidth * 3.5);
		assert.strictEqual(para.GetText({ParaEndToSpace : false}), "1234", "Paragraph text: 1234");
		assert.strictEqual(para.GetLinesCount(), 2, "Lines count 2");
		assert.strictEqual(para.GetTextOnLine(0), "123", "Text on line 0 '123'");
		assert.strictEqual(para.GetTextOnLine(1), "4", "Text on line 1 '4'");

		SetText("12 34");
		Recalculate(charWidth * 3.5);
		assert.strictEqual(para.GetText({ParaEndToSpace : false}), "12 34", "Paragraph text: 12 34");
		assert.strictEqual(para.GetLinesCount(), 2, "Lines count 2");
		assert.strictEqual(para.GetTextOnLine(0), "12 ", "Text on line 0 '12 ");
		assert.strictEqual(para.GetTextOnLine(1), "34", "Text on line 1 '34'");

	});

	QUnit.test("Test: \"Test paragraph with very narrow width\"", function (assert)
	{
		assert.strictEqual(dc.GetElementsCount(), 1, "Paragraphs count");

		let narrowWidth = charWidth / 2;

		SetText("");
		Recalculate(narrowWidth);
		assert.strictEqual(para.GetLinesCount(), 1, "Lines count of empty paragraph");
		assert.deepEqual(para.GetLineBounds(0), new AscWord.CDocumentBounds(0, 0, narrowWidth, AscTest.FontHeight), "Line bounds of empty paragraph");
		assert.strictEqual(para.GetPagesCount(), 1, "Pages count of paragraph");

		SetText("123");
		Recalculate(narrowWidth);
		assert.strictEqual(para.GetText({ParaEndToSpace : false}), "123", "Paragraph text: 123");
		assert.strictEqual(para.GetLinesCount(), 3, "Lines count 3");
		assert.deepEqual(para.GetLineBounds(0), new AscWord.CDocumentBounds(0, 0, narrowWidth, AscTest.FontHeight), "Check line bounds");
		assert.deepEqual(para.GetLineBounds(1), new AscWord.CDocumentBounds(0, AscTest.FontHeight, narrowWidth, AscTest.FontHeight * 2), "Check line bounds");
		assert.deepEqual(para.GetLineBounds(2), new AscWord.CDocumentBounds(0, AscTest.FontHeight * 2, narrowWidth, AscTest.FontHeight * 3), "Check line bounds");
		assert.strictEqual(para.GetPagesCount(), 1, "Pages count of paragraph");

		SetText("Q");
		Recalculate(narrowWidth);
		assert.strictEqual(para.GetText({ParaEndToSpace : false}), "Q", "Paragraph text: Q");
		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.deepEqual(para.GetLineBounds(0), new AscWord.CDocumentBounds(0, 0, narrowWidth, AscTest.FontHeight), "Check line bounds");
		assert.strictEqual(para.GetPagesCount(), 1, "Pages count of paragraph");

		SetText("Q ");
		Recalculate(narrowWidth);
		assert.strictEqual(para.GetText({ParaEndToSpace : false}), "Q ", "Paragraph text: Q'<'space'>'");
		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.deepEqual(para.GetLineBounds(0), new AscWord.CDocumentBounds(0, 0, narrowWidth, AscTest.FontHeight), "Check line bounds");
		assert.strictEqual(para.GetPagesCount(), 1, "Pages count of paragraph");
	});

	QUnit.test("Test: \"Test line break of ligatures\"", function (assert)
	{
		SetText("ffi");

		let text_f_0 = run.GetElement(0);
		let text_f_1 = run.GetElement(1);
		let text_i   = run.GetElement(2);

		Recalculate(charWidth * 3.5);
		assert.strictEqual(para.GetText({ParaEndToSpace : false}), "ffi", "Paragraph text: ffi");

		assert.strictEqual(text_f_0.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE, "Check f code point type");
		assert.strictEqual(text_f_1.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE, "Check f code point type");
		assert.strictEqual(text_i.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE, "Check i code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "ffi", "Text on line 0 'ffi ");

		Recalculate(charWidth * 2.5);

		assert.strictEqual(text_f_0.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE, "Check f code point type");
		assert.strictEqual(text_f_1.GetCodePointType(), AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE, "Check f code point type");
		assert.strictEqual(text_i.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check i code point type");

		assert.strictEqual(para.GetLinesCount(), 2, "Lines count 2");
		assert.strictEqual(para.GetTextOnLine(0), "ff", "Text on line 0 'ff'");
		assert.strictEqual(para.GetTextOnLine(1), "i", "Text on line 1 'i'");

		Recalculate(charWidth * 1.5);

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
		SetText("xyz");

		let text_x = run.GetElement(0);
		let text_y = run.GetElement(1);
		let text_z = run.GetElement(2);

		Recalculate(charWidth * 3.5);
		assert.strictEqual(para.GetText({ParaEndToSpace : false}), "xyz", "Paragraph text: xyz");

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz ");

		// Комбинированные символы НЕ ДОЛЖНЫ отдельно переносится на новую строку
		Recalculate(charWidth * 2.5);

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz'");

		Recalculate(charWidth * 1.5);

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz'");

		Recalculate(charWidth * 1.5);

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz'");

		Recalculate(charWidth * 0.5);

		assert.strictEqual(text_x.GetCodePointType(), AscWord.CODEPOINT_TYPE.BASE, "Check x code point type");
		assert.strictEqual(text_y.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check y code point type");
		assert.strictEqual(text_z.GetCodePointType(), AscWord.CODEPOINT_TYPE.COMBINING_MARK, "Check z code point type");

		assert.strictEqual(para.GetLinesCount(), 1, "Lines count 1");
		assert.strictEqual(para.GetTextOnLine(0), "xyz", "Text on line 0 'xyz'");
	});


});
