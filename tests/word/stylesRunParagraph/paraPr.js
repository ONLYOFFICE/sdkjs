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


$(function () {

	let logicDocument = AscTest.CreateLogicDocument();
	let stylesManager = logicDocument.GetStyles();

	let strNameFirstParagraph	= "FirstParagraph";
	let strNameSecondParagraph	= "SecondParagraph";
	let strNameFirstParaRun = "FirstParaRun";
	let strNameSecondParaRun = "SecondParaRun";

	let oFirstParagraphStyle = new AscWord.CStyle(strNameFirstParagraph, null, null, styletype_Paragraph);
	let oSecondParagraphStyle = new AscWord.CStyle(strNameSecondParagraph, null, null, styletype_Paragraph);

	let oFirstParaRunStyle = new AscWord.CStyle(strNameFirstParaRun, null, null, styletype_Character);
	let oSecondParaRunStyle = new AscWord.CStyle(strNameSecondParaRun, null, null, styletype_Character);

	stylesManager.Add(oFirstParagraphStyle);
	stylesManager.Add(oSecondParagraphStyle);
	stylesManager.Add(oFirstParaRunStyle);
	stylesManager.Add(oSecondParaRunStyle);

	function GetText()
	{
		logicDocument.SelectAll();
		let text = logicDocument.GetSelectedText();
		logicDocument.RemoveSelection();
		return text;
	}

	function StartTest(text, styleParagraph, styleRun)
	{
		AscTest.ClearDocument();
		let p = new AscWord.CParagraph(AscTest.DrawingDocument);
		p.SetParagraphStyleById(styleParagraph);
		logicDocument.AddToContent(0, p);

		p.Add_ToContent(0, Run(p, text, styleRun))
		logicDocument.MoveCursorToEndPos();
	}

	function Run(parent, text, style)
	{
		let r = new AscWord.CRun(parent);
		r.SetRStyle(style);
		r.AddText(text);
		return r;
	}

	QUnit.module("Paragraph style (ParaPr)");

	QUnit.test("Indents", function (assert)
	{
		StartTest("World!", oFirstParagraphStyle.GetId(), oFirstParaRunStyle.GetId());

		AscTest.MoveCursorLeft(true, false, 7); // World!

		assert.strictEqual(logicDocument.GetSelectedText(false), "World!", "Add text 'Hello World!'");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParaRunStyle.Name, "If select style Run in style Paragraph -> RunStyle");

		console.log(logicDocument.GetStyleFromFormatting());
	});
});
