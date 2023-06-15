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
	
	let oFirstParagraphStyle  = AscTest.CreateParagraphStyle("ParaStyle1");
	let oSecondParagraphStyle = AscTest.CreateParagraphStyle("ParaStyle2");
	
	let oFirstParaRunStyle = AscTest.CreateRunStyle("RunStyle1");
	let oSecondParaRunStyle = AscTest.CreateRunStyle("RunStyle2");

	function AddParagraph(pos, style)
	{
		let p = AscTest.CreateParagraph();
		
		if (style)
			p.SetParagraphStyleById(style.GetId());
		
		logicDocument.AddToContent(pos, p);
		return p;
	}
	function CreateRun(text, style)
	{
		let r = AscTest.CreateRun();

		if (style)
			r.SetRStyle(style.GetId());

		r.AddText(text);
		return r;
	}

	QUnit.module("Cursor");

	QUnit.test("Run with style in paragraph", function (assert)
	{
		let word = "Word!";

		AscTest.ClearDocument();
		let p = AddParagraph(0, oFirstParagraphStyle);
		p.AddToContent(0, CreateRun(word, oFirstParaRunStyle));
		p.MoveCursorToEndPos();

		AscTest.MoveCursorLeft(false, false, 1);
		assert.ok(true, "Move cursor to the last letter of the word");
		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParaRunStyle.Name, "Must show run style");
	});

	QUnit.test("Run without style in paragraph", function (assert)
	{
		let word = "Word!";

		AscTest.ClearDocument();
		let p = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, CreateRun(word));
		assert.ok(true, "Create run without style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(false, false, word.length - (word.length - 1));
		assert.ok(true, "Move cursor to first letter of word");

		assert.strictEqual(logicDocument.GetStyleFromFormatting().BasedOn, oFirstParagraphStyle.Name, "AddParagraph style");
	});

	QUnit.module("Select");

	QUnit.test("Add two paragraphs with different styles", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, CreateRun(strFirstWord));
		assert.ok(true, "Create run without style");
		let p2 = AddParagraph(1, oSecondParagraphStyle);
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, CreateRun(strSecondWord));
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
		let p1 = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, CreateRun(strFirstWord, oFirstParaRunStyle));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		let p2 = AddParagraph(1, oSecondParagraphStyle);
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, CreateRun(strSecondWord, oSecondParaRunStyle));
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
		let p1 = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, CreateRun(strFirstWord));
		assert.ok(true, "Create run without style");
		let p2 = AddParagraph(1, oSecondParagraphStyle);
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, CreateRun(strSecondWord));
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
		let p = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, CreateRun(strFirstWord, oFirstParaRunStyle));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p.Add_ToContent(1, CreateRun(strSecondWord, oFirstParaRunStyle));
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
		let p = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, CreateRun(strFirstWord, oFirstParaRunStyle));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p.Add_ToContent(1, CreateRun(strSecondWord, oFirstParaRunStyle));
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
		let p = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p.Add_ToContent(0, CreateRun(strFirstWord, oFirstParaRunStyle));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p.Add_ToContent(1, CreateRun(strSecondWord, oSecondParaRunStyle));
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
		let p1 = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, CreateRun(strFirstWord, oFirstParaRunStyle));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p1.Add_ToContent(1, CreateRun(strDel, oSecondParaRunStyle));
		assert.ok(true, "Create another run with oSecondParaRunStyle style");

		let p2 = AddParagraph(1, oSecondParagraphStyle);
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, CreateRun(strSecondWord, oSecondParaRunStyle));
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
		let p1 = AddParagraph(0, oFirstParagraphStyle);
		assert.ok(true, "Create paragraph with oFirstParagraphStyle style");
		p1.Add_ToContent(0, CreateRun(strFirstWord, oFirstParaRunStyle));
		assert.ok(true, "Create run with oFirstParaRunStyle style");
		p1.Add_ToContent(1, CreateRun(strDel, oSecondParaRunStyle));
		assert.ok(true, "Create another run with oSecondParaRunStyle style with spaces and tabs");

		let p2 = AddParagraph(1, oSecondParagraphStyle);
		assert.ok(true, "Create paragraph with oSecondParagraphStyle style");
		p2.Add_ToContent(0, CreateRun(strSecondWord, oFirstParaRunStyle));
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
