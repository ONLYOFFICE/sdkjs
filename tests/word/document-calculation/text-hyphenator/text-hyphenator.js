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
	const charWidth = AscTest.CharWidth * AscTest.FontSize;
	
	let dc = new AscWord.DocumentContent();
	dc.ClearContent(false);
	
	let para = new AscWord.Paragraph();
	dc.AddToContent(0, para);
	
	let run = new AscWord.Run();
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
	
	let autoHyphenation = false;
	let hyphenateCaps   = true;
	let hyphenLimit     = 0;
	
	AscWord.Paragraph.prototype.isAutoHyphenation = function()
	{
		return autoHyphenation;
	};
	AscWord.ParagraphRecalculationWrapState.prototype.isAutoHyphenation = function()
	{
		return autoHyphenation;
	};
	AscWord.ParagraphRecalculationWrapState.prototype.getAutoHyphenLimit = function()
	{
		return hyphenLimit;
	};
	AscWord.TextHyphenator.prototype.isHyphenateCaps = function()
	{
		return hyphenateCaps;
	};
	
	function setAutoHyphenation(isAuto)
	{
		autoHyphenation = isAuto;
	}
	function setHyphenateCaps(isHyphenate)
	{
		hyphenateCaps = isHyphenate;
	}
	function setHyphenLimit(limit)
	{
		hyphenLimit = limit;
	}
	
	function checkLines(assert, isAutoHyphenation, contentWidth, textLines)
	{
		setAutoHyphenation(isAutoHyphenation);
		recalculate(contentWidth);
		
		assert.strictEqual(para.GetLinesCount(), textLines.length, "Check lines count " + textLines.length);
		
		for (let line = 0, lineBreakPos = 0; line < textLines.length; ++line)
		{
			lineBreakPos += textLines[line].length;
			
			let lineText = textLines[line];
			if (textLines[line].length && "-" === textLines[line].charAt(textLines[line].length - 1))
			{
				--lineBreakPos;
				checkAutoHyphenAfter(assert, lineBreakPos - 1, true);
				lineText = textLines[line].substr(0, textLines[line].length - 1);
			}
			else
			{
				checkAutoHyphenAfter(assert, lineBreakPos - 1, false);
			}
			
			assert.strictEqual(para.GetTextOnLine(line), lineText, "Text on line " + line + " '" + textLines[line] + "'");
		}
	}
	
	function checkAutoHyphenAfter(assert, itemPos, isHyphen, _run)
	{
		let __run = _run ? _run : run;
		let runItem = __run.GetElement(itemPos);
		if (!runItem.IsText())
			assert.strictEqual(false, isHyphen, "Check auto hyphen after symbol");
		else
			assert.strictEqual(runItem.IsTemporaryHyphenAfter(), isHyphen, "Check auto hyphen after symbol");
	}
	
	QUnit.module("Text hyphenation",
	{
		beforeEach : function ()
		{
			setAutoHyphenation(false);
			setHyphenateCaps(true);
			setHyphenLimit(0);
		}
	});
	
	QUnit.test("Test: \"Test regular line break cases\"", function (assert)
	{
		setText("abcd abcd aaabbb");
		checkLines(assert, false, charWidth * 8.5, [
			"abcd ",
			"abcd ",
			"aaabbb"
		]);
		checkLines(assert, true, charWidth * 8.5, [
			"abcd ab-",
			"cd aaa-",
			"bbb"
		]);
		// Дефис переноса не убирается
		checkLines(assert, true, charWidth * 7.5, [
			"abcd ",
			"abcd ",
			"aaabbb"
		]);
		
		// Перенос на первой букве
		setText("abbb");
		checkLines(assert, false, charWidth * 3.5, [
			"abb",
			"b",
		]);
		checkLines(assert, true, charWidth * 3.5, [
			"a-",
			"bbb",
		]);
		
		setText("aabbbcccdddd");
		checkLines(assert, false, charWidth * 3.5, [
			"aab",
			"bbc",
			"ccd",
			"ddd"
		]);
		checkLines(assert, true, charWidth * 3.5, [
			"aa-",
			"bbb",
			"ccc",
			"ddd",
			"d"
		]);
	});
	
	QUnit.test("Test: \"Test edge cases\"", function (assert)
	{
		setText("aaaa zz½www bbbb");
		checkLines(assert, false, charWidth * 7.5, [
			"aaaa ",
			"zz½www ",
			"bbbb"
		]);
		checkLines(assert, true, charWidth * 7.5, [
			"aaaa ",
			"zz½www ",
			"bbbb"
		]);
		checkLines(assert, true, charWidth * 8.5, [
			"aaaa zz-",
			"½www bbbb"
		]);

		// Перенос идет после второго символа z, а следующий за ним символ меньше по ширине, чем
		// размер дефиса, который мы рисуем во время переноса
		setText("zz½www");
		checkLines(assert, false, charWidth * 2.75, [
			"zz½",
			"ww",
			"w"
		]);
		checkLines(assert, true, charWidth * 3.25, [
			"zz-",
			"½ww",
			"w"
		]);
		checkLines(assert, true, charWidth * 2.75, [
			"zz½",
			"ww",
			"w"
		]);
		checkLines(assert, true, charWidth * 2.25, [
			"zz",
			"½w",
			"ww"
		]);
		
		// TODO: Разобрать случай, когда перенос слова происходит в двух (или более местах) и следующее место переноса
		//       надо начинать считать с последнего места переноса, а не с начала слова
		
		// TODO: Случай, когда одно длинное слово разбивается по переносам и целиком переходит на следующую страницу
		//       из-за этого
	});
	
	QUnit.test("Test: \"Test DoNotHyphenateCaps parameter\"", function (assert)
	{
		setText("abcde AAABBB aaabbb");
		
		checkLines(assert, false, charWidth * 11.5, [
			"abcde ",
			"AAABBB ",
			"aaabbb"
		]);
		
		setHyphenateCaps(true);
		checkLines(assert, true, charWidth * 11.5, [
			"abcde AAA-",
			"BBB aaabbb"
		]);
		
		setHyphenateCaps(false);
		checkLines(assert, true, charWidth * 11.5, [
			"abcde ",
			"AAABBB aaa-",
			"bbb"]
		);
	});
	
	QUnit.test("Test: \"Test ConsecutiveHyphenLimit parameter for different words\"", function (assert)
	{
		setText("abcd AAABBBB aaabbbb aaabbbb AAABBBB aaabbbb aaabbbb abcd");
		setText("abcd AAABBB aaabbb aaabbb AAABBB aaabbb aaabbb abcd");
		
		checkLines(assert, false, charWidth * 9.5, [
			"abcd ",
			"AAABBB ",
			"aaabbb ",
			"aaabbb ",
			"AAABBB ",
			"aaabbb ",
			"aaabbb ",
			"abcd"]
		);
		checkAutoHyphenAfter(assert, 7, false);
		checkAutoHyphenAfter(assert, 14, false);
		checkAutoHyphenAfter(assert, 21, false);
		checkAutoHyphenAfter(assert, 28, false);
		checkAutoHyphenAfter(assert, 35, false);
		checkAutoHyphenAfter(assert, 42, false);
		
		// checkLines(assert, true, charWidth * 9.5, [
		// 	"abcd AAA",
		// 	"BBB aaa",
		// 	"bbb aaa",
		// 	"bbb AAA",
		// 	"BBB aaa",
		// 	"bbb aaa",
		// 	"bbb abcd"]
		// );
		// checkAutoHyphenAfter(assert, 7, true);
		// checkAutoHyphenAfter(assert, 14, true);
		// checkAutoHyphenAfter(assert, 21, true);
		// checkAutoHyphenAfter(assert, 28, true);
		// checkAutoHyphenAfter(assert, 35, true);
		// checkAutoHyphenAfter(assert, 42, true);
		//
		// setHyphenLimit(1);
		// checkLines(assert, true, charWidth * 10.5, [
		// 	"abcd AAA",
		// 	"BBBB ",
		// 	"aaabbb aaa",
		// 	"bbb ",
		// 	"AAABBB aaa",
		// 	"bbb ",
		// 	"aaabbb abcd"]
		// );
		// checkAutoHyphenAfter(assert, 7, true);
		// checkAutoHyphenAfter(assert, 14, false);
		// checkAutoHyphenAfter(assert, 21, true);
		// checkAutoHyphenAfter(assert, 28, false);
		// checkAutoHyphenAfter(assert, 35, true);
		// checkAutoHyphenAfter(assert, 42, false);
		//
		// setHyphenLimit(2);
		// checkLines(assert, true, charWidth * 10.5, [
		// 	"abcd AAA",
		// 	"BBB aaa",
		// 	"bbb ",
		// 	"aaabbb ",
		// 	"AAABBB aaa",
		// 	"bbb aaa",
		// 	"bbb abcd"]
		// );
		// checkAutoHyphenAfter(assert, 7, true);
		// checkAutoHyphenAfter(assert, 14, false);
		// checkAutoHyphenAfter(assert, 21, true);
		// checkAutoHyphenAfter(assert, 28, false);
		// checkAutoHyphenAfter(assert, 35, true);
		// checkAutoHyphenAfter(assert, 42, false);
		//
		// setHyphenLimit(3);
		// checkLines(assert, true, charWidth * 10.5, [
		// 	"abcd AAA",
		// 	"BBB aaa",
		// 	"bbb aaa",
		// 	"bbb ",
		// 	"AAABBB aaa",
		// 	"bbb aaa",
		// 	"bbb abcd"]
		// );
		// checkAutoHyphenAfter(assert, 7, true);
		// checkAutoHyphenAfter(assert, 14, true);
		// checkAutoHyphenAfter(assert, 21, true);
		// checkAutoHyphenAfter(assert, 28, false);
		// checkAutoHyphenAfter(assert, 35, true);
		// checkAutoHyphenAfter(assert, 42, true);
		
		// setText("abcd aabbb ABBBB abbb AABBB abbbb aaabbbb abcd");
		//
		// checkLines(assert, false, charWidth * 8.5, [
		// 	"abcd ",
		// 	"aabbb ",
		// 	"ABBBB ",
		// 	"abbb ",
		// 	"AABBB ",
		// 	"abbbb ",
		// 	"aaabbbb ",
		// 	"abcd"]
		// );
		//
		// checkLines(assert, true, charWidth * 8.5, [
		// 	"abcd aa-",
		// 	"bbb A-",
		// 	"BBBB a-",
		// 	"bbb AA-",
		// 	"BBB a-",
		// 	"bbbb aaa-",
		// 	"bbbb ",
		// 	"abcd"]
		// );
	});
	
	/*
	QUnit.test("Test: \"Test ConsecutiveHyphenLimit parameter for single word\"", function (assert)
	{
		setText("aabbbcccdddd");
		
		checkLines(assert, false, charWidth * 4.5, [
			"aabb",
			"bccc",
			"dddd"
		]);
		
		checkLines(assert, true, charWidth * 4.5, [
			"aa-",
			"bbb-",
			"ccc-",
			"dddd"
		]);
		
		// В этом примере важно, что ccdddd тоже переносится по второму символу
		setHyphenLimit(1);
		checkLines(assert, true, charWidth * 4.5, [
			"aa-",
			"bbbc",
			"cc-",
			"dddd"
		]);
		
		setHyphenLimit(2);
		checkLines(assert, true, charWidth * 4.5, [
			"aa-",
			"bbb-",
			"cccd",
			"ddd"
		]);
		
	});
	*/
});
