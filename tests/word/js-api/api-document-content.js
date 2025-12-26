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

$(function ()
{
	QUnit.module("ApiDocumentContent");
	
	
	QUnit.test("GetText", function (assert)
	{
		let docContent = AscTest.JsApi.CreateDocContent();
		let p = docContent.GetElement(0);
		let run = p.AddText("123");
		run.AddTabStop();
		run.AddText("456");
		run.AddLineBreak();
		run.AddText("789");
		
		let table = AscTest.JsApi.CreateTable(2,2);
		table.GetRow(0).GetCell(0).GetContent().GetElement(0).AddText("A");
		table.GetRow(0).GetCell(1).GetContent().GetElement(0).AddText("B");
		table.GetRow(1).GetCell(0).GetContent().GetElement(0).AddText("C");
		table.GetRow(1).GetCell(1).GetContent().GetElement(0).AddText("D");
		
		docContent.Push(table);
		
		assert.strictEqual(docContent.GetText(), "123\t456\r789\r\nA\tB\r\nC\tD\r\n\r\n", "Check GetText");
		assert.strictEqual(docContent.GetText({
			"TabSymbol" : "_t_",
			"NewLineSeparator" : "_nl_",
			"TableCellSeparator" : "_c_",
			"TableRowSeparator" : "_r_",
			"ParaSeparator" : "_p_"
			
		}), "123_t_456_nl_789_p_A_c_B_r_C_c_D_r__p_", "Check GetText");
	});
});
