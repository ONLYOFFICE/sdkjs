/*
 * (c) Copyright Ascensio System SIA 2010-2026
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
	const logicDocument = AscTest.CreateLogicDocument();
	const eps = 0.01;

	function setupDocument()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(400, 400);
		sectPr.SetPageMargins(50, 50, 50, 50);
	}
	
	function createTableNoCellW(rows, cols, tableGrid)
	{
		let table = AscTest.CreateTable(rows, cols);
		table.TableGrid = tableGrid;
		return table;
	}

	QUnit.module("Fixed layout", {
		beforeEach : function()
		{
			setupDocument();
		}
	});

	QUnit.test("Test: cell widths smaller than grid shrink columns", function (assert)
	{
		let table = createTableNoCellW(2, 3, [50, 50, 50]);
		table.SetTableLayout(tbllayout_Fixed);
		logicDocument.PushToContent(table);

		table.GetRow(0).GetCell(0).SetW(new CTableMeasurement(tblwidth_Mm, 30));
		table.GetRow(0).GetCell(1).SetW(new CTableMeasurement(tblwidth_Mm, 25));
		table.GetRow(0).GetCell(2).SetW(new CTableMeasurement(tblwidth_Mm, 40));

		AscTest.Recalculate();

		assert.close(table.TableGridCalc[0], 30, eps, "Column 0 should shrink from 50 to 30");
		assert.close(table.TableGridCalc[1], 25, eps, "Column 1 should shrink from 50 to 25");
		assert.close(table.TableGridCalc[2], 40, eps, "Column 2 should shrink from 50 to 40");
		
		AscTest.ClearDocument();
		table = AscTest.CreateTable(2, 3, [50, 50, 50]);
		table.SetTableLayout(tbllayout_Fixed);
		logicDocument.PushToContent(table);
		
		table.GetRow(0).GetCell(0).SetW(new CTableMeasurement(tblwidth_Mm, 30));
		table.GetRow(0).GetCell(1).SetW(new CTableMeasurement(tblwidth_Mm, 25));
		table.GetRow(0).GetCell(2).SetW(new CTableMeasurement(tblwidth_Mm, 40));
		
		AscTest.Recalculate();
		
		assert.close(table.TableGridCalc[0], 50, eps, "Column 0 should shrink from 50 to 30");
		assert.close(table.TableGridCalc[1], 50, eps, "Column 1 should shrink from 50 to 25");
		assert.close(table.TableGridCalc[2], 50, eps, "Column 2 should shrink from 50 to 40");
	});

	QUnit.test("Test: cell widths larger than grid expand columns", function (assert)
	{
		let table = createTableNoCellW(1, 3, [30, 30, 30]);
		table.SetTableLayout(tbllayout_Fixed);
		logicDocument.PushToContent(table);

		table.GetRow(0).GetCell(0).SetW(new CTableMeasurement(tblwidth_Mm, 50));
		table.GetRow(0).GetCell(1).SetW(new CTableMeasurement(tblwidth_Mm, 60));
		table.GetRow(0).GetCell(2).SetW(new CTableMeasurement(tblwidth_Mm, 70));

		AscTest.Recalculate();

		assert.close(table.TableGridCalc[0], 50, eps, "Column 0 should expand from 30 to 50");
		assert.close(table.TableGridCalc[1], 60, eps, "Column 1 should expand from 30 to 60");
		assert.close(table.TableGridCalc[2], 70, eps, "Column 2 should expand from 30 to 70");
	});

	QUnit.test("Test: maximum cell width wins across multiple rows", function (assert)
	{
		let table = createTableNoCellW(3, 2, [50, 50]);
		table.SetTableLayout(tbllayout_Fixed);
		logicDocument.PushToContent(table);

		table.GetRow(0).GetCell(0).SetW(new CTableMeasurement(tblwidth_Mm, 30));
		table.GetRow(1).GetCell(0).SetW(new CTableMeasurement(tblwidth_Mm, 45));
		table.GetRow(2).GetCell(0).SetW(new CTableMeasurement(tblwidth_Mm, 35));

		table.GetRow(0).GetCell(1).SetW(new CTableMeasurement(tblwidth_Mm, 60));
		table.GetRow(1).GetCell(1).SetW(new CTableMeasurement(tblwidth_Mm, 55));
		table.GetRow(2).GetCell(1).SetW(new CTableMeasurement(tblwidth_Mm, 58));

		AscTest.Recalculate();

		assert.close(table.TableGridCalc[0], 45, eps, "Column 0 should be max(30, 45, 35) = 45");
		assert.close(table.TableGridCalc[1], 60, eps, "Column 1 should be max(60, 55, 58) = 60");
	});

	QUnit.test("Test: auto width cells fall back to original grid", function (assert)
	{
		let table = createTableNoCellW(1, 3, [50, 60, 70]);
		table.SetTableLayout(tbllayout_Fixed);
		logicDocument.PushToContent(table);

		AscTest.Recalculate();

		assert.close(table.TableGridCalc[0], 50, eps, "Column 0 should keep original grid value 50");
		assert.close(table.TableGridCalc[1], 60, eps, "Column 1 should keep original grid value 60");
		assert.close(table.TableGridCalc[2], 70, eps, "Column 2 should keep original grid value 70");
	});

	QUnit.test("Test: mixed explicit and auto widths", function (assert)
	{
		let table = createTableNoCellW(2, 3, [50, 60, 70]);
		table.SetTableLayout(tbllayout_Fixed);
		logicDocument.PushToContent(table);

		table.GetRow(0).GetCell(1).SetW(new CTableMeasurement(tblwidth_Mm, 40));

		AscTest.Recalculate();

		assert.close(table.TableGridCalc[0], 50, eps, "Column 0 (no explicit width) should keep original grid 50");
		assert.close(table.TableGridCalc[1], 40, eps, "Column 1 should use cell width 40");
		assert.close(table.TableGridCalc[2], 70, eps, "Column 2 (no explicit width) should keep original grid 70");
	});

	QUnit.test("Test: vmerge_Continue cells are ignored", function (assert)
	{
		let table = createTableNoCellW(2, 2, [50, 50]);
		table.SetTableLayout(tbllayout_Fixed);
		logicDocument.PushToContent(table);

		table.GetRow(0).GetCell(0).SetW(new CTableMeasurement(tblwidth_Mm, 30));
		table.GetRow(1).GetCell(0).SetVMerge(vmerge_Continue);
		table.GetRow(1).GetCell(0).SetW(new CTableMeasurement(tblwidth_Mm, 80));

		AscTest.Recalculate();

		assert.close(table.TableGridCalc[0], 30, eps, "Column 0 should be 30 (vmerge_Continue cell with 80 is ignored)");
	});
});
