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
	// Temporary we use epsilon equal to 2 twips instead of 1
	const epsilon = AscCommon.TwipsToMM(2);
	
	const logicDocument = AscTest.CreateLogicDocument();


	QUnit.module("Test the positioning of flow tables");

	QUnit.test("Flow table inside another table", function (assert)
	{
		// Здесь мы проверяем горизонтальную позицию вложенной таблицы

		logicDocument.Settings.CompatibilityMode = AscCommon.document_compatibility_mode_Word15;
		AscTest.ClearDocument();

		let p1 = AscTest.CreateParagraph();
		logicDocument.PushToContent(p1);

		let table = AscTest.CreateTable(1, 1);
		logicDocument.PushToContent(table);
		AscTest.RemoveTableBorders(table);

		let cell = table.GetRow(0).GetCell(0);
		cell.Set_Margins({
			Left   : {W : 10, Type : tblwidth_Mm},
			Top    : {W : 0, Type : tblwidth_Mm},
			Right  : {W : 10, Type : tblwidth_Mm},
			Bottom : {W : 0, Type : tblwidth_Mm}
		}, -1);

		let p2 = AscTest.CreateParagraph();
		logicDocument.PushToContent(p2);

		let sectPr = logicDocument.GetLastSection();
		sectPr.SetPageMargins(20, 20, 20, 20);

		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Check pages count of the document");
		assert.strictEqual(table.GetPageBounds(0).Left, 20, "Check left position of outer table");

		let innerTable  = AscTest.CreateTable(1, 1);
		let cellContent = cell.GetContent();
		cellContent.PushToContent(innerTable);

		AscTest.RemoveTableBorders(innerTable);

		AscTest.Recalculate();
		assert.strictEqual(innerTable.GetPageBounds(0).Left, 30, "Check left position of the inner inline table");

		innerTable.SetInline(false);
		innerTable.SetPositionH(Asc.c_oAscHAnchor.Page, false, 0);
		innerTable.SetPositionV(Asc.c_oAscVAnchor.Page, false, 0);

		AscTest.Recalculate();
		assert.strictEqual(innerTable.GetPageBounds(0).Left, 20, "Check left position of the inner flow table (Page position 0)");

		// https://bugzilla.onlyoffice.com/show_bug.cgi?id=46314
		innerTable.SetPositionH(Asc.c_oAscHAnchor.Margin, false, 0);
		AscTest.Recalculate();
		assert.strictEqual(innerTable.GetPageBounds(0).Left, 30, "Check left position of the inner inline table (Margin position 0)");

		innerTable.SetPositionH(Asc.c_oAscHAnchor.Margin, true, Asc.c_oAscXAlign.Left);
		AscTest.Recalculate();
		assert.strictEqual(innerTable.GetPageBounds(0).Left, 30, "Check left position of the inner inline table (Margin align left)");

	});

	QUnit.test("Check table position in second section on page", function (assert)
	{
		// https://bugzilla.onlyoffice.com/show_bug.cgi?id=42163

		logicDocument.Settings.CompatibilityMode = AscCommon.document_compatibility_mode_Word15;
		AscTest.ClearDocument();

		let p1 = AscTest.CreateParagraph();
		logicDocument.PushToContent(p1);

		let table = AscTest.CreateTable(1, 1);
		logicDocument.PushToContent(table);
		AscTest.RemoveTableBorders(table);

		let p2 = AscTest.CreateParagraph();
		logicDocument.PushToContent(p2);

		let sectPr = logicDocument.GetLastSection();
		sectPr.SetType(Asc.c_oAscSectionBreakType.Continuous);
		sectPr.SetPageMargins(20, 20, 20, 20);

		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Check pages count of the document");
		assert.strictEqual(table.GetPageBounds(0).Left, 20, "Check the left position of the inline table");

		sectPr = new AscWord.CSectionPr();
		sectPr.SetPageMargins(40, 40, 40, 40);
		sectPr.SetType(Asc.c_oAscSectionBreakType.Continuous);
		p1.SetSectionPr(sectPr);

		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Check pages count of the document");
		assert.strictEqual(table.GetPageBounds(0).Left, 20, "Add section to the page and check the left position of the inline table");

		p1.RemoveSectionPr();

		table.SetInline(false);
		table.SetPositionH(Asc.c_oAscHAnchor.Margin, true, Asc.c_oAscXAlign.Left);
		table.SetPositionV(Asc.c_oAscVAnchor.Page, false, 100);

		AscTest.Recalculate();
		assert.ok(!table.IsInline(), "Make table float");
		assert.strictEqual(table.GetPageBounds(0).Left, 20, "Check the left position of the float table");

		p1.SetSectionPr(sectPr);

		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Check pages count of the document");
		assert.strictEqual(table.GetPageBounds(0).Left, 20, "Add section to the page and check the left position of the float table");
	});
	
	QUnit.test("Check table overlap", function (assert)
	{
		AscTest.ClearDocument();

		let sectPr = logicDocument.GetLastSection();
		sectPr.SetPageMargins(20, 20, 20, 20);
		
		const rowHeight = 30;
		function createTable(rows, cols)
		{
			let t = AscTest.CreateTable(rows, cols);
			for (let iRow = 0; iRow < t.GetRowsCount(); ++iRow)
			{
				t.GetRow(iRow).SetHeight(rowHeight, Asc.linerule_AtLeast);
			}
			
			t.SetInline(false);
			t.SetPositionH(Asc.c_oAscHAnchor.Page, false, 0);
			t.SetPositionV(Asc.c_oAscVAnchor.Page, false, 0);
			AscTest.RemoveTableBorders(t);
			return t;
		}
		
		let table1 = createTable(2, 2);
		logicDocument.AddToContent(0, table1);
		
		let table2 = createTable(2, 2);
		logicDocument.AddToContent(1, table2);
		
		table1.setAllowOverlap(true);
		table2.setAllowOverlap(true);
		
		table1.SetPositionV(Asc.c_oAscVAnchor.Page, false, 0);
		table2.SetPositionV(Asc.c_oAscVAnchor.Page, false, 10);

		AscTest.Recalculate();
		assert.close(table1.GetPageBounds(0).Top, 0, epsilon, "Check top position of the first table");
		assert.close(table2.GetPageBounds(0).Top, 10, epsilon, "Check top position of the second table");
		
		table1.setAllowOverlap(false);
		table2.setAllowOverlap(true);
		
		AscTest.Recalculate();
		assert.close(table1.GetPageBounds(0).Top, 0, epsilon, "Check top position of the first table");
		assert.close(table2.GetPageBounds(0).Top, 2 * rowHeight, epsilon, "Check top position of the second table");
		
		table1.SetPositionH(Asc.c_oAscHAnchor.Page, false, 50);
		table2.SetPositionH(Asc.c_oAscHAnchor.Page, false, 50);

		table1.setAllowOverlap(true);
		table2.setAllowOverlap(true);
		
		table1.SetPositionV(Asc.c_oAscVAnchor.Text, false, 0);
		table2.SetPositionV(Asc.c_oAscVAnchor.Text, false, 10);
		
		AscTest.Recalculate();
		assert.close(table1.GetPageBounds(0).Top, 20, epsilon, "Check top position of the first table");
		assert.close(table2.GetPageBounds(0).Top, 30, epsilon, "Check top position of the second table");
		
		table1.setAllowOverlap(false);
		table2.setAllowOverlap(false);
		
		AscTest.Recalculate();
		assert.close(table1.GetPageBounds(0).Top, 20, epsilon, "Check top position of the first table");
		assert.close(table2.GetPageBounds(0).Top, 20 + 2 * rowHeight, epsilon, "Check top position of the second table");
	});
	
	
	QUnit.test("Check flow table positioning in a header/footer (bug 67673)", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = AscTest.CreateParagraph();
		logicDocument.PushToContent(p1);
		
		let sectPr = logicDocument.GetLastSection();
		sectPr.SetPageMargins(50, 50, 50, 50);
		sectPr.SetPageSize(100, 200);
		sectPr.SetPageMarginHeader(20);
		
		let t = AscTest.CreateTable(1, 1);
		t.GetRow(0).SetHeight(150, Asc.linerule_AtLeast);
		t.SetInline(false);
		t.SetPositionH(Asc.c_oAscHAnchor.Page, false, 0);
		t.SetPositionV(Asc.c_oAscVAnchor.Page, false, 20);
		AscTest.RemoveTableBorders(t);
		
		let header = AscTest.CreateDefaultHeader(sectPr);
		header.AddToContent(1, t);
		
		p1 = header.GetElement(0);
		p1.SetParagraphSpacing({Before : 0, After : 0, Line : 1, LineRule : Asc.linerule_Auto});
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		AscTest.Recalculate();
		assert.strictEqual(t.GetPagesCount(), 1, "Check page count");
		assert.strictEqual(t.IsEmptyPage(0), false, "Check page content");
		assert.close(t.GetPageBounds(0).Top, 20, epsilon, "Check top position of the table");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word14);
		AscTest.Recalculate();
		assert.strictEqual(t.GetPagesCount(), 1, "Check top position of the first table");
		assert.strictEqual(t.IsEmptyPage(0), false, "Check page content");
		assert.close(t.GetPageBounds(0).Top, 20, epsilon, "Check top position of the table");
		
		t.SetPositionV(Asc.c_oAscVAnchor.Text, false, 0);

		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word14);
		AscTest.Recalculate();
		assert.strictEqual(t.GetPagesCount(), 2, "Check page count");
		assert.strictEqual(t.IsEmptyPage(0), true, "Check page content");
		assert.strictEqual(t.IsEmptyPage(1), false, "Check page content");
		assert.close(t.GetPageBounds(0).Top, 20 + AscTest.FontHeight, epsilon, "Check top position of the table");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		AscTest.Recalculate();
		assert.strictEqual(t.GetPagesCount(), 1, "Check page count");
		assert.strictEqual(t.IsEmptyPage(0), false, "Check page content");
		assert.close(t.GetPageBounds(0).Top, 20 + AscTest.FontHeight, epsilon, "Check top position of the table");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
	});
});
