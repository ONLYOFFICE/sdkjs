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

$(function ()
{
	// For correct calculation of a tab position we need to create logic document
	// since position depends on overall document margins
	let logicDocument = AscTest.CreateLogicDocument();
	let charWidth     = AscTest.CharWidth * AscTest.FontSize;
	
	let sectPr = AscTest.GetFinalSection();
	sectPr.SetPageSize(100 * charWidth, 1000);
	sectPr.SetPageMargins(10 * charWidth, 50, 15 * charWidth, 50);
	
	let p = AscTest.CreateParagraph();
	logicDocument.AddToContent(0, p);
	
	let r = AscTest.CreateRun();
	p.AddToContent(0, r);
	
	function setTabs(tabs)
	{
		let paraTabs = new CParaTabs();
		tabs.forEach(t => paraTabs.Add(new CParaTab(t.value, t.pos, t.leader)));
		p.SetParagraphTabs(paraTabs);
	}
	
	QUnit.module("Paragraph tabs calculation");
	
	QUnit.test("Special case for left tab which exceed right edge", function (assert)
	{
		// Check situation when left tab lies between right edge of a paragraph and right field of the document
		r.AddText("Before\tafter");
		p.SetParagraphIndent({Right : 20 * charWidth});
		setTabs([{value : tab_Left, pos : 70 * charWidth}]);
		
		AscTest.Recalculate();
		assert.strictEqual(p.GetLinesCount(), 1, "Check number of lines");
		assert.strictEqual(p.GetTextOnLine(0), "Before after", "Text on line 0 'Before after'");
	});
});
