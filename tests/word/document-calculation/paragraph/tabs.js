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


$(function ()
{
	let charWidth    = AscTest.CharWidth * AscTest.FontSize;
	let contentWidth = 20 * charWidth;
	
	let dc = new AscWord.CDocumentContent();
	dc.ClearContent(false);
	let p = new AscWord.CParagraph();
	dc.AddToContent(0, p);
	
	let r = new AscWord.CRun();
	p.AddToContent(0, r);
	
	function recalculate()
	{
		dc.Reset(0, 0, contentWidth, 10000);
		dc.Recalculate_Page(0, true);
	}
	
	function setTabs(tabs)
	{
		let paraTabs = new CParaTabs();
		tabs.forEach(t => paraTabs.Add(new CParaTab(t.value, t.pos, t.leader)));
		p.SetParagraphTabs(paraTabs);
	}
	
	QUnit.module("Paragraph tabs calculation");
	
	QUnit.test("Special case for left tab which exceed right edge", function (assert)
	{
		r.AddText("Before\tafter");
		setTabs([{value : tab_Left, pos : contentWidth + 5 * charWidth}]);
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word14);
		recalculate();
		assert.strictEqual(p.GetLinesCount(), 1, "Check number of lines");
		assert.strictEqual(p.GetTextOnLine(0), "Before after", "Text on line 0 'Before after'");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		recalculate();
		assert.strictEqual(p.GetLinesCount(), 2, "Check number of lines");
		assert.strictEqual(p.GetTextOnLine(0), "Before", "Text on line 0 'Before'");
		assert.strictEqual(p.GetTextOnLine(1), " after", "Text on line 0 'after'");

	});
});
