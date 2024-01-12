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
	const logicDocument = AscTest.CreateLogicDocument()
	
	function setupDocument()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(400, 400);
		sectPr.SetPageMargins(50, 50, 50, 50);
	}
	
	
	QUnit.module("Test various situations with table header", {
		beforeEach : function()
		{
			setupDocument();
		}
	});
	
	QUnit.test("Test page break", function (assert)
	{
		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		paragraph.SetParagraphSpacing({Before: 150});
		
		AscTest.Recalculate();
		assert.deepEqual(paragraph.GetPageBounds(0), new AscWord.CDocumentBounds(50, 50, 350, 200 + AscTest.FontHeight), "Check page bounds of the first paragraph");
		
		let table = AscTest.CreateTable(4, 3, [100, 100, 100]);
		logicDocument.PushToContent(table);

		AscTest.RemoveTableBorders(table);

		table.GetRow(0).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(1).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(2).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(3).SetHeight(50, Asc.linerule_AtLeast);

		// Test a normal table divided into two pages
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "Check pages count");
		assert.deepEqual(table.getRowBounds(2, 1), new AscWord.CDocumentBounds(50, 50, 350, 100), "Check row bounds of the first row on the second page");
		
		// Test table with 1 heading row divided into two pages
		table.GetRow(0).SetHeader(true);
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "Check pages count");
		assert.deepEqual(table.getRowBounds(2, 1), new AscWord.CDocumentBounds(50, 100, 350, 150), "Check row bounds of the first row on the second page");
		
		// Test case when no regular rows on the first page and whole table should start from the second page (#62031)
		paragraph.SetParagraphSpacing({Before: 200});
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "Check pages count");
		assert.strictEqual(table.IsEmptyPage(0), true, "First page should be empty");
		assert.strictEqual(table.IsEmptyPage(1), false, "Check second page");
	});
	
});
