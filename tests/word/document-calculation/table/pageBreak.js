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
		sectPr.SetPageSize(400, 820);
		sectPr.SetPageMargins(10, 10, 10, 10);
	}
	
	
	QUnit.module("Test various situations with breaking table on pages", {
		beforeEach : function()
		{
			setupDocument();
		}
	});
	
	QUnit.test("Test page break because of bottom border width", function (assert)
	{
		let table = AscTest.CreateTable(4, 3);
		logicDocument.PushToContent(table);
		
		AscTest.RemoveTableBorders(table);
		
		table.GetRow(0).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(1).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(2).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(3).SetHeight(50, Asc.linerule_AtLeast);
		
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 1, "Should be 1 page");
		
		table.GetRow(0).SetHeight(700, Asc.linerule_AtLeast);
		table.GetRow(1).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(2).SetHeight(45, Asc.linerule_AtLeast);
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "Should be 2 pages");
		assert.strictEqual(table.GetPage(1).GetFirstRow(), 3, "Check that page breaks on the 4-th row");
		
		table.SelectRows(2, 2);
		table.Set_Props({
			CellBorders : {Bottom : AscWord.CBorder.FromObject({Value : border_Single, Size : 10, Space : 0})}
		});
		table.RemoveSelection();
		AscTest.Recalculate();
		assert.ok(true, "Increase the thickness of the bottom border of the second row");
		assert.strictEqual(table.GetPagesCount(), 2, "Should be 2 pages");
		assert.strictEqual(table.GetPage(1).GetFirstRow(), 2, "Check that page breaks on the 3-th row");
	});
	
	QUnit.test("Test page break of a float table in a special case (bug 57159)", function (assert)
	{
		// Таблица имеет прилегание по вертикали к параграфу с некоторым смещением и не убирается целиком на странице
		let table = AscTest.CreateTable(8, 3);
		logicDocument.AddToContent(0, table);
		AscTest.RemoveTableBorders(table);
		
		table.SetInline(false);
		table.SetPositionH(Asc.c_oAscHAnchor.Page, false, 0);
		table.SetPositionV(Asc.c_oAscVAnchor.Text, false, 0);
		
		for (let iRow = 0, nRows = table.GetRowsCount(); iRow < nRows; ++iRow)
		{
			table.GetRow(iRow).SetHeight(50, Asc.linerule_AtLeast);
		}
		
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 1, "Should be 1 page");
		
		table.SetPositionV(Asc.c_oAscVAnchor.Text, false, 605);
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "Should be 1 page");
		assert.strictEqual(table.GetPage(1).GetFirstRow(), 3, "Check that page breaks on the 4-th row");
		
		table.SetPositionV(Asc.c_oAscVAnchor.Text, false, 505);
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "Should be 1 page");
		assert.strictEqual(table.GetPage(1).GetFirstRow(), 5, "Check that page breaks on the 4-th row");
	});
	
});
