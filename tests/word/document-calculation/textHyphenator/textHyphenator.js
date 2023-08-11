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
	
	let dc = new AscWord.CDocumentContent();
	dc.ClearContent(false);
	
	let para = new AscWord.CParagraph();
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
	
	let autoHyphenation = false;
	let hyphenateCaps   = true;
	
	AscWord.CParagraph.prototype.IsAutoHyphenation = function()
	{
		return autoHyphenation;
	};
	AscWord.CTextHyphenator.prototype.IsHyphenateCaps = function()
	{
		return hyphenateCaps;
	};
	
	function SetAutoHyphenation(isAuto)
	{
		autoHyphenation = isAuto;
	}
	function SetHyphenateCaps(isHyphenate)
	{
		hyphenateCaps = isHyphenate;
	}
	
	function CheckLines(assert, isAutoHyphenation, contentWidth, textLines)
	{
		SetAutoHyphenation(isAutoHyphenation);
		Recalculate(contentWidth);
		
		assert.strictEqual(para.GetLinesCount(), textLines.length, "Check lines count " + textLines.length);
		
		for (let line = 0; line < textLines.length; ++line)
		{
			assert.strictEqual(para.GetTextOnLine(line), textLines[line], "Text on line " + line + " '" + textLines[line] + "'");
		}
	}
	
	function CheckAutoHyphenAfter(assert, itemPos, isHyphen, _run)
	{
		let __run = _run ? _run : run;
		assert.strictEqual(__run.GetElement(itemPos).IsTemporaryHyphenAfter(), isHyphen, "Check auto hyphen after symbol");
	}
	
	QUnit.module("Paragraph Lines");
	
	QUnit.test("Test: \"Test regular line break cases\"", function (assert)
	{
		SetText("abcd abcd aaabbb");
		CheckLines(assert, false, charWidth * 8.5, ["abcd ", "abcd ", "aaabbb"]);
		CheckAutoHyphenAfter(assert, 6, false);
		CheckLines(assert, true, charWidth * 8.5, ["abcd ab", "cd aaa", "bbb"]);
		CheckAutoHyphenAfter(assert, 6, true);
		// Дефис переноса не убирается
		CheckLines(assert, true, charWidth * 7.5, ["abcd ", "abcd ", "aaabbb"]);
		CheckAutoHyphenAfter(assert, 6, false);
		
		SetText("aabbbcccdddd");
		CheckLines(assert, false, charWidth * 3.5, ["aab", "bbc", "ccd", "ddd"]);
		CheckLines(assert, true, charWidth * 3.5, ["aa", "bbb", "ccc", "ddd", "d"]);
	});
	
	QUnit.test("Test: \"Test edge cases\"", function (assert)
	{
		SetText("aaaa zz½www bbbb");
		CheckLines(assert, false, charWidth * 7.5, ["aaaa ", "zz½www ", "bbbb"]);
		CheckLines(assert, true, charWidth * 7.5, ["aaaa ", "zz½www ", "bbbb"]);
		CheckLines(assert, true, charWidth * 8.5, ["aaaa zz", "½www bbbb"]);

		// Перенос идет после второго символа z, а следующий за ним символ меньше по ширине, чем
		// размер дефиса, который мы рисуем во время переноса
		SetText("zz½www");
		CheckLines(assert, false, charWidth * 2.75, ["zz½", "ww", "w"]);
		CheckAutoHyphenAfter(assert, 1, false);
		CheckLines(assert, true, charWidth * 3.25, ["zz", "½ww", "w"]);
		CheckAutoHyphenAfter(assert, 1, true);
		CheckLines(assert, true, charWidth * 2.75, ["zz½", "ww", "w"]);
		CheckAutoHyphenAfter(assert, 1, false);
		CheckLines(assert, true, charWidth * 2.25, ["zz", "½w", "ww"]);
		CheckAutoHyphenAfter(assert, 1, false);
	});
	
	QUnit.test("Test: \"Test DoNotHyphenateCaps parameter\"", function (assert)
	{
		SetText("abcde AAABBB aaabbb");
		
		CheckLines(assert, false, charWidth * 11.5, ["abcde ", "AAABBB ", "aaabbb"]);
		CheckAutoHyphenAfter(assert, 8, false);
		CheckAutoHyphenAfter(assert, 15, false);
		
		SetHyphenateCaps(true);
		CheckLines(assert, true, charWidth * 11.5, ["abcde AAA", "BBB aaabbb"]);
		CheckAutoHyphenAfter(assert, 8, true);
		CheckAutoHyphenAfter(assert, 15, false);
		
		SetHyphenateCaps(false);
		CheckLines(assert, true, charWidth * 11.5, ["abcde ", "AAABBB aaa", "bbb"]);
		CheckAutoHyphenAfter(assert, 8, false);
		CheckAutoHyphenAfter(assert, 15, true);
		
	});

});
