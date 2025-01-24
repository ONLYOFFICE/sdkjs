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
	const lField = 50;
	const rField = 350;
	
	function checkBounds(actual, expected, message)
	{
		let maxDifference = 0.01;
		
		QUnit.assert.close(actual.Left, expected.Left, maxDifference, "Left: " + message);
		QUnit.assert.close(actual.Top, expected.Top, maxDifference, "Top: " + message);
		QUnit.assert.close(actual.Right, expected.Right, maxDifference, "Right: " + message);
		QUnit.assert.close(actual.Bottom, expected.Bottom, maxDifference, "Bottom: " + message);
	}
	
	const logicDocument = AscTest.CreateLogicDocument()
	
	function setupDocument()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(400, 400);
		sectPr.SetPageMargins(50, 50, 50, 50);
	}
	
	function addImageToParagraph(p, x, y, w, h)
	{
		let d = AscTest.CreateImage(w, h);
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		run.AddToContent(0, d);
		
		d.Set_DrawingType(drawing_Anchor);
		d.Set_Distance(10, 10, 10, 10);
		d.Set_WrappingType(WRAPPING_TYPE_SQUARE);
		d.Set_PositionH(Asc.c_oAscRelativeFromH.Page, false, x);
		d.Set_PositionV(Asc.c_oAscRelativeFromV.Page, false, y);
		
		return d;
	}
	
	
	QUnit.module("Test various situations with table header", {
		beforeEach : function()
		{
			setupDocument();
			AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
		}
	});
	
	QUnit.test("Single flow-image on the left side", function(assert)
	{
		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		paragraph.SetParagraphSpacing({Before : 0, After : 0});
		
		let table = AscTest.CreateTable(3, 3, [100, 100, 100]);
		logicDocument.PushToContent(table);
		AscTest.RemoveTableBorders(table);
		AscTest.RemoveTableMargins(table);
		
		table.GetRow(0).SetHeight(20, Asc.linerule_AtLeast);
		table.GetRow(1).SetHeight(20, Asc.linerule_AtLeast);
		table.GetRow(2).SetHeight(20, Asc.linerule_AtLeast);
		
		let tableTop = 50 + AscTest.FontHeight;
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		
		// Test a normal situation
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 1, "Check table page count");
		checkBounds(table.getRowBounds(0, 0), new AscWord.CDocumentBounds(lField, tableTop, rField, tableTop + 20), "Check first row bounds");
		checkBounds(table.getRowBounds(1, 0), new AscWord.CDocumentBounds(lField, tableTop + 20, rField, tableTop + 40), "Check second row bounds");
		checkBounds(table.getRowBounds(2, 0), new AscWord.CDocumentBounds(lField, tableTop + 40, rField, tableTop + 60), "Check third row bounds");
		
		addImageToParagraph(paragraph, lField, tableTop, 50, 50);
		
		// AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		//
		// AscTest.Recalculate();
		// assert.strictEqual(table.GetPagesCount(), 1, "Check table page count");
		// checkBounds(table.getRowBounds(0, 0), new AscWord.CDocumentBounds(lField, tableTop + 60, rField, tableTop + 80), "Check first row bounds");
		// checkBounds(table.getRowBounds(1, 0), new AscWord.CDocumentBounds(lField, tableTop + 80, rField, tableTop + 100), "Check second row bounds");
		// checkBounds(table.getRowBounds(2, 0), new AscWord.CDocumentBounds(lField, tableTop + 100, rField, tableTop + 120), "Check third row bounds");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word14);
		
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 1, "Check table page count");
		let _rField = rField + table.GetRightTableOffsetCorrection() - table.GetTableOffsetCorrection();
		checkBounds(table.getRowBounds(0, 0), new AscWord.CDocumentBounds(lField + 60, tableTop, _rField, tableTop + 20), "Check first row bounds");
		checkBounds(table.getRowBounds(1, 0), new AscWord.CDocumentBounds(lField + 60, tableTop + 20, _rField, tableTop + 40), "Check second row bounds");
		checkBounds(table.getRowBounds(2, 0), new AscWord.CDocumentBounds(lField + 60, tableTop + 40, _rField, tableTop + 60), "Check third row bounds");
	});
	
});
