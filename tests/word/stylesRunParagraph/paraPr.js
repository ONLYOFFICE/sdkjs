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

	function Paragraph(pos, style)
	{
		let p = new AscWord.CParagraph(AscTest.DrawingDocument);

		if (style)
			p.SetParagraphStyleById(style);

		logicDocument.AddToContent(pos, p);

		return p;
	}
	function Run(parent, text, style)
	{
		let r = new AscWord.CRun(parent);

		if (style)
			r.SetRStyle(style);

		r.AddText(text);
		return r;
	}

	QUnit.module("Cursor");

	QUnit.test("Run with style in paragraph", function (assert)
	{
		let strWord = "Word!";

		AscTest.ClearDocument();
		let p = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, Run(p, strWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(false, false, strWord.length - (strWord.length - 1));
		assert.ok(true, "Move cursor to first letter of word");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParaRunStyle.Name, "Must show run style");
	});

	QUnit.test("Run without style in paragraph", function (assert)
	{
		let word = "Word!";

		AscTest.ClearDocument();
		let p = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, Run(p, word));
		assert.ok(true, "Create run without style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(false, false, word.length - (word.length - 1));
		assert.ok(true, "Move cursor to first letter of word");

		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParagraphStyle.Name, "Paragraph style");
	});

	QUnit.module("Select");

	QUnit.test("Add two paragraphs with different styles", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, Run(p1, strFirstWord));
		assert.ok(true, "Create run without style");
		let p2 = Paragraph(1, oSecondParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, Run(p2, strSecondWord));
		assert.ok(true, "Create run without style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord.length);
		assert.ok(true, "Select two runs");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParagraphStyle.Name, "If different paragraphs are selected, the style of the first one is visible - oFirstParagraphStyle");
	});

	QUnit.test("Select part of run with style", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, Run(p1, strFirstWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		let p2 = Paragraph(1, oSecondParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, Run(p2, strSecondWord, oSecondParaRunStyle.GetId()));
		assert.ok(true, "Create run with oSecondParaRunStyle style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length  - 1);
		assert.ok(true, "Select part of first run");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oSecondParaRunStyle.Name, "Must show style of run - oSecondParaRunStyle");
	});

	QUnit.test("Select part of paragraph with style", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, Run(p1, strFirstWord));
		assert.ok(true, "Create run without style");
		let p2 = Paragraph(1, oSecondParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, Run(p2, strSecondWord));
		assert.ok(true, "Create run without style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length  - 1);
		assert.ok(true, "Select part of Paragraph");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oSecondParagraphStyle.Name, "Must show style of second selected paragraph - oSecondParagraphStyle");

		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length );
		assert.ok(true, "Select all second paragraph");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oSecondParagraphStyle.Name, "Must show style of second selected paragraph - oSecondParagraphStyle");

		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord);
		assert.ok(true, "Select first and second paragraph");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParagraphStyle.Name, "Must show style of first selected paragraph - oFirstParagraphStyle");
	});

	QUnit.test("Runs with same style in one paragraph", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, Run(p, strFirstWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p.Add_ToContent(1, Run(p, strSecondWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length);
		assert.ok(true, "Select second Run");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParaRunStyle.Name, "Must show style of run - oFirstParaRunStyle");

		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord.length);
		assert.ok(true, "Select all runs");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParaRunStyle.Name, "Must show style of runs - oFirstParaRunStyle");
	});

	QUnit.test("Multiple runs with same style in one paragraph", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, Run(p, strFirstWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p.Add_ToContent(1, Run(p, strSecondWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create another run with oFirstParaRunStyle style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord.length);
		assert.ok(true, "Select all runs");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParaRunStyle.Name, "Must show style of runs oFirstParaRunStyle");
	});

	QUnit.test("Multiple runs with different style in one paragraph", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, Run(p, strFirstWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p.Add_ToContent(1, Run(p, strSecondWord, oSecondParaRunStyle.GetId()));
		assert.ok(true, "Create run with oSecondParaRunStyle style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length);
		assert.ok(true, "Select second run");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oSecondParaRunStyle.Name, "Must show style of last run - oSecondParaRunStyle");

		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord.length);
		assert.ok(true, "Select second run");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParagraphStyle.Name, "Must show style of paragraph, so runs with different style - oFirstParagraphStyle");
	});

	QUnit.test("Multiple runs with same style in two different style paragraphs", function (assert)
	{
		let strFirstWord = "Word!";
		let strDel = "Del"
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, Run(p1, strFirstWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p1.Add_ToContent(1, Run(p1, strDel, oSecondParaRunStyle.GetId()));
		assert.ok(true, "Create another run with oSecondParaRunStyle style");

		let p2 = Paragraph(1, oSecondParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, Run(p2, strSecondWord, oSecondParaRunStyle.GetId()));
		assert.ok(true, "Create run with oSecondParaRunStyle style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strDel.length);
		assert.ok(true, "Select two last runs");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oSecondParaRunStyle.Name, "Must show style of two last runs - oSecondParaRunStyle");

		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strDel.length + strFirstWord.length);
		assert.ok(true, "Select all runs");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParagraphStyle.Name, "Must show style of first Paragraph, because style of runs in selection are different - oSecondParaRunStyle");
	});

	QUnit.test("Space anf tab in run", function (assert)
	{
		let strFirstWord = "Word!";
		let strDel = "   	   "
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = Paragraph(0, oFirstParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, Run(p1, strFirstWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p1.Add_ToContent(1, Run(p1, strDel, oSecondParaRunStyle.GetId()));
		assert.ok(true, "Create another run with oSecondParaRunStyle style with spaces and tabs");

		let p2 = Paragraph(1, oSecondParagraphStyle.GetId());
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, Run(p2, strSecondWord, oFirstParaRunStyle.GetId()));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strDel.length);
		assert.ok(true, "Select two last runs");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParaRunStyle.Name, "Must show style of last run, runs with spaces and other style dont affect - oFirstParaRunStyle");

		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strDel.length + strFirstWord.length);
		assert.ok(true, "Select all runs");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParaRunStyle.Name, "Must show style of last run, runs with spaces and other style dont affect - oFirstParaRunStyle");
	});
});