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


$(function () {

	let dc = new AscWord.CDocumentContent();
	dc.ClearContent(false);

	let p1 = new AscWord.Paragraph();
	let p2 = new AscWord.Paragraph();

	dc.AddToContent(0, p1);
	dc.AddToContent(1, p2);

	let r1 = new AscWord.CRun();
	p1.AddToContent(0, r1);
	r1.AddText("Hello Word!");

	let r2 = new AscWord.CRun();
	p2.AddToContent(0, r2);
	r2.AddText("Абракадабра");

	const pageWidth = 20 * AscTest.CharWidth * AscTest.FontSize;
	function Recalculate()
	{
		dc.Reset(0, 0, 20 * AscTest.CharWidth * AscTest.FontSize, 10000);
		dc.Recalculate_Page(0, true);
	}

	QUnit.module("Paragraph Spacing");


	QUnit.test("Test: \"Paragraphs\"", function (assert)
	{
		p1.SetParagraphSpacing({Before : 0, After : 0});
		p2.SetParagraphSpacing({Before : 0, After : 0});

		Recalculate();
		assert.strictEqual(dc.GetElementsCount(), 2, "Check paragraphs count");
		assert.strictEqual(p1.GetPagesCount(), 1, "Check pages count of the first paragraph");
		assert.strictEqual(p2.GetPagesCount(), 1, "Check pages count of the second paragraph");

		assert.deepEqual(p1.GetPageBounds(0), new AscWord.CDocumentBounds(0, 0, pageWidth, AscTest.FontHeight), "Check page bounds of the first paragraph");
		assert.deepEqual(p2.GetPageBounds(0), new AscWord.CDocumentBounds(0, AscTest.FontHeight, pageWidth, AscTest.FontHeight * 2), "Check page bounds of the second paragraph");


		p1.SetParagraphSpacing({Before : 15, After : 20});
		p2.SetParagraphSpacing({Before : 0, After : 0});

		Recalculate();
		assert.deepEqual(p1.GetPageBounds(0), new AscWord.CDocumentBounds(0, 0, pageWidth, AscTest.FontHeight + 35), "Check page bounds of the first paragraph");
		assert.deepEqual(p2.GetPageBounds(0), new AscWord.CDocumentBounds(0, AscTest.FontHeight + 35, pageWidth, AscTest.FontHeight * 2 + 35), "Check page bounds of the second paragraph");

		p1.SetParagraphSpacing({Before : 15, After : 20});
		p2.SetParagraphSpacing({Before : 30, After : 0});

		Recalculate();
		assert.deepEqual(p1.GetPageBounds(0), new AscWord.CDocumentBounds(0, 0, pageWidth, AscTest.FontHeight + 35), "Check page bounds of the first paragraph");
		assert.deepEqual(p2.GetPageBounds(0), new AscWord.CDocumentBounds(0, AscTest.FontHeight + 35, pageWidth, AscTest.FontHeight * 2 + 45), "Check page bounds of the second paragraph");
	});
});
