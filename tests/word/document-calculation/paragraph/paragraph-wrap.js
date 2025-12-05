/*
 * (c) Copyright Ascensio System SIA 2010-2025
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
	
	let logicDocument = AscTest.CreateLogicDocument();
	
	const charWidth = AscTest.CharWidth * AscTest.FontSize;
	
	const L_FIELD = 20 * charWidth;
	const R_FIELD = 30 * charWidth;
	const PAGE_W  = 150 * charWidth;

	
	function addFlowImageToParagraph(p, w, h, x, y)
	{
		let d = AscTest.CreateImage(w, h);
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		run.AddToContent(0, d);
		
		d.Set_DrawingType(drawing_Anchor);
		d.Set_Distance(0, 0, 0, 0);
		d.Set_WrappingType(WRAPPING_TYPE_SQUARE);
		d.Set_PositionH(Asc.c_oAscRelativeFromH.Page, false, x);
		d.Set_PositionV(Asc.c_oAscRelativeFromV.Page, false, y);
		return d;
	}
	
	function initDocument()
	{
		AscTest.ClearDocument();
		logicDocument.AddToContent(0, AscTest.CreateParagraph());
		
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(PAGE_W, 1000);
		sectPr.SetPageMargins(L_FIELD, 50, R_FIELD, 50);
	}
	
	function checkText(assert, para, lines)
	{
		assert.strictEqual(para.getLineCount(), lines.length, "Check number of lines");
		
		for (let i = 0, lineCount = Math.min(lines.length, para.GetLinesCount()); i < lineCount; ++i)
		{
			assert.strictEqual(para.getRangeCount(i), lines[i].length, "Check number of ranges in " + (i + 1) + " line");
			for (let j = 0, rangeCount = Math.min(lines[i].length, para.getRangeCount(i)); j < rangeCount; ++j)
			{
				assert.strictEqual(para.getTextInLineRange(i, j), lines[i][j], "Check text in range " + j);
			}
		}
	}
	function checkRangeBounds(assert, para, lines)
	{
		for (let i = 0, lineCount = Math.min(lines.length, para.GetLinesCount()); i < lineCount; ++i)
		{
			for (let j = 0, rangeCount = Math.min(lines[i].length, para.getRangeCount(i)); j < rangeCount; ++j)
			{
				let range = para.getRange(i, j);
				assert.close(range.X, lines[i][j][0], 0.001, "Check x for " + j + "range");
				assert.close(range.XEnd, lines[i][j][1], 0.001, "Check x for " + j + "range");
			}
		}
	}
	
	QUnit.module("Test paragraph wrap", {
		beforeEach : function()
		{
			initDocument();
			AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
		}
	});
	
	QUnit.test("Test simple wrap", function(assert)
	{
		let imageX0 = 45 * charWidth;
		let imageX1 = imageX0 + 30 * charWidth;
		let imageY0 = 40;
		let imageY1 = imageY0 + 50;
		
		let p1 = logicDocument.GetElement(0);
		
		let p2 = AscTest.CreateParagraph();
		
		logicDocument.AddToContent(0, p2);
		
		addFlowImageToParagraph(p2, imageX1 - imageX0, imageY1 - imageY0, imageX0, imageY0);
		
		AscTest.AddTextToParagraph(p1, "VeryLongWord The quick brown fox jumps over the lazy dog");
		AscTest.Recalculate();
		
		checkText(assert, p1, [
			["VeryLongWord The quick ", "brown fox jumps over the lazy dog\r\n"]
		]);
		checkRangeBounds(assert, p1, [
			[[L_FIELD, imageX0], [imageX1, PAGE_W - R_FIELD]]
		]);
		
		// Check ranges order for RTL paragraph
		p1.SetParagraphBidi(true);
		AscTest.Recalculate();
		checkText(assert, p1, [
			["VeryLongWord The quick brown fox jumps over ", "the lazy dog\r\n"]
		]);
		checkRangeBounds(assert, p1, [
			[[imageX1, PAGE_W - R_FIELD], [L_FIELD, imageX0]]
		]);
	});
	
	QUnit.test("Test the first line indent", function(assert)
	{
		let imageX0   = 45 * charWidth;
		let imageX1   = imageX0 + 10 * charWidth;
		let imageY0   = 40;
		let imageY1   = 90;
		let firstLine = 5 * charWidth;
		let leftInd   = 5 * charWidth;
		let rightInd  = 0;
		
		let text = "VeryLongWord The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog.";
		
		function initParagraphWithImage()
		{
			initDocument();
			
			let p1 = logicDocument.GetElement(0);
			let p2 = AscTest.CreateParagraph();
			logicDocument.AddToContent(0, p2);
			
			addFlowImageToParagraph(p2, imageX1 - imageX0, imageY1 - imageY0, imageX0, imageY0);
			
			p1.SetParagraphIndent({FirstLine : firstLine, Left : leftInd, Right : rightInd});
			p1.SetParagraphSpacing({Line : 1, LineRule : Asc.linerule_Auto, Before : 0, After : 0});
			AscTest.AddTextToParagraph(p1, text);
			AscTest.Recalculate();
			return p1;
		}
		
		function test(textLines, ranges)
		{
			let p = initParagraphWithImage();
			checkText(assert, p, textLines);
			checkRangeBounds(assert, p, ranges);
		}
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		
		// Simple situation
		test([
			["VeryLongWord ", "The quick brown fox jumps over the lazy dog. The quick brown fox "],
			["jumps over the lazy ", "dog.\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1, PAGE_W - R_FIELD]],
			[[L_FIELD + leftInd, imageX0], [imageX1, PAGE_W - R_FIELD]]
		]);
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word14);
		
		text = "VeryLongWord The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog.";
		
		firstLine = 5 * charWidth;
		leftInd   = 15 * charWidth;
		test([
			["", "VeryLongWord The quick brown fox jumps over the lazy dog. "],
			["The quick ", "brown fox jumps over the lazy dog. The quick brown fox jumps over "],
			["the lazy dog.\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + firstLine, PAGE_W - R_FIELD]],
			[[L_FIELD + leftInd, imageX0], [imageX1, PAGE_W - R_FIELD]],
			[[L_FIELD + leftInd, PAGE_W - R_FIELD]]
		]);
		
		firstLine = 15 * charWidth;
		leftInd   = 0;
		test([
			["", "VeryLongWord The quick brown fox jumps over the "],
			["lazy dog. The quick brown ", "fox jumps over the lazy dog. The quick brown fox jumps over the "],
			["lazy dog.\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + firstLine, PAGE_W - R_FIELD]],
			[[L_FIELD + leftInd, imageX0], [imageX1, PAGE_W - R_FIELD]],
			[[L_FIELD + leftInd, PAGE_W - R_FIELD]]
		]);
		
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		
		firstLine = 5 * charWidth;
		leftInd   = 15 * charWidth;
		test([
			["", "VeryLongWord The quick brown fox jumps over the lazy dog. "],
			["The quick ", "brown fox jumps over the lazy dog. The quick brown fox jumps over "],
			["the lazy dog.\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + firstLine, PAGE_W - R_FIELD]],
			[[L_FIELD + leftInd, imageX0], [imageX1, PAGE_W - R_FIELD]],
			[[L_FIELD + leftInd, PAGE_W - R_FIELD]]
		]);
		
		
		// Check the indentation when the first range is empty
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word14);
		
		firstLine = 0;
		leftInd   = 0;
		text      = "VeryLongLongLongLongLongLongLongWord"
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + firstLine, PAGE_W - R_FIELD]]
		]);
		
		firstLine = 10 * charWidth;
		leftInd   = 0;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + firstLine, PAGE_W - R_FIELD]]
		]);
		
		firstLine = 10 * charWidth;
		leftInd   = 10 * charWidth;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + firstLine, PAGE_W - R_FIELD]]
		]);
		
		firstLine = -5 * charWidth;
		leftInd   = 10 * charWidth;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + leftInd + firstLine, PAGE_W - R_FIELD]]
		]);
		
		firstLine = -15 * charWidth;
		leftInd   = 10 * charWidth;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + leftInd + firstLine, PAGE_W - R_FIELD]]
		]);
		
		firstLine = -5 * charWidth;
		leftInd   = 10 * charWidth;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + leftInd + firstLine, PAGE_W - R_FIELD]]
		]);
		
		firstLine = 0;
		leftInd   = 10 * charWidth;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1, PAGE_W - R_FIELD]]
		]);

		// Check the indentation when the first range is empty
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		
		firstLine = 0;
		leftInd   = 0;
		text      = "VeryLongLongLongLongLongLongLongWord"
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + Math.max(0, firstLine), PAGE_W - R_FIELD]]
		]);
		
		firstLine = 10 * charWidth;
		leftInd   = 0;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + Math.max(0, firstLine), PAGE_W - R_FIELD]]
		]);
		
		firstLine = 10 * charWidth;
		leftInd   = 10 * charWidth;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + Math.max(0, firstLine), PAGE_W - R_FIELD]]
		]);
		
		firstLine = -5 * charWidth;
		leftInd   = 10 * charWidth;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + Math.max(0, firstLine), PAGE_W - R_FIELD]]
		]);
		
		firstLine = -15 * charWidth;
		leftInd   = 10 * charWidth;
		test([
			["", "VeryLongLongLongLongLongLongLongWord\r\n"],
		], [
			[[L_FIELD + firstLine + leftInd, imageX0], [imageX1 + Math.max(0, firstLine), PAGE_W - R_FIELD]]
		]);
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
		
	});
});
