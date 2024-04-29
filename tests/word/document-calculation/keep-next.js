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
	const logicDocument = AscTest.CreateLogicDocument()
	
	function setupDocument()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(400, 400);
		sectPr.SetPageMargins(50, 50, 50, 50);
	}
	
	
	QUnit.module("Test various situations keepNext property", {
		beforeEach : function()
		{
			setupDocument();
		}
	});
	
	QUnit.test("Test a simple situation with three paragraphs: keepNext = [false, true, false]", function (assert)
	{
		let p1 = AscTest.CreateParagraph();
		let p2 = AscTest.CreateParagraph();
		let p3 = AscTest.CreateParagraph();
		
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		logicDocument.PushToContent(p3);
		
		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Should be 1 page");
		
		p1.SetParagraphSpacing({After: 250});

		AscTest.Recalculate();
		assert.strictEqual(p2.GetPagesCount(), 1, "p2 have 1 page");
		assert.strictEqual(p3.GetPagesCount(), 2, "p3 have 2 pages");
		assert.strictEqual(p3.IsEmptyPage(0) && !p3.IsEmptyPage(1), true, "p3 the first page is empty and the second is not");
		
		p2.SetParagraphKeepNext(true);
		AscTest.Recalculate();
		assert.strictEqual(p2.GetPagesCount(), 2, "p2 have 2 pages");
		assert.strictEqual(p2.IsEmptyPage(0) && !p2.IsEmptyPage(1), true, "p2 the first page is empty and the second is not");
		assert.strictEqual(p3.GetPagesCount(), 1, "p3 have 1 page");
		assert.strictEqual(p3.GetStartPageAbsolute(), 1, "check p3 start page number");
	});
	
	QUnit.test("Test the case when a paragraph with the KeepNext property is followed by a table", function (assert)
	{
		let p1 = AscTest.CreateParagraph();
		let p2 = AscTest.CreateParagraph();
		
		let table = AscTest.CreateTable(3, 3);
		table.GetRow(0).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(1).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(2).SetHeight(50, Asc.linerule_AtLeast);
		
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		logicDocument.PushToContent(table);
		
		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Should be 1 page");
		
		p1.SetParagraphSpacing({After: 250});
		
		AscTest.Recalculate();
		assert.strictEqual(p2.GetPagesCount(), 1, "p2 have 1 page");
		assert.strictEqual(table.GetPagesCount(), 2, "table have 2 pages");
		assert.strictEqual(table.IsEmptyPage(0) && !table.IsEmptyPage(1), true, "table's the first page is empty and the second is not");

		p2.SetParagraphKeepNext(true);
		AscTest.Recalculate();
		assert.strictEqual(p2.GetPagesCount(), 2, "p2 have 2 pages");
		assert.strictEqual(p2.IsEmptyPage(0) && !p2.IsEmptyPage(1), true, "p2 the first page is empty and the second is not");
		assert.strictEqual(table.GetPagesCount(), 1, "table have 1 page");
		assert.strictEqual(table.GetStartPageAbsolute(), 1, "check table start page number");
	});
	
});
