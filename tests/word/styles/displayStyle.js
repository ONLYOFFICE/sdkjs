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
	let styleManager  = logicDocument.GetStyleManager();
	
	let paraStyle1 = AscTest.CreateParagraphStyle("ParaStyle1");
	let paraStyle2 = AscTest.CreateParagraphStyle("ParaStyle2");
	
	let runStyle1 = AscTest.CreateRunStyle("RunStyle1");
	let runStyle2 = AscTest.CreateRunStyle("RunStyle2");
	
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
	
	let displayStyleName = "";
	editor.isDocumentLoadComplete = true;
	editor.UpdateParagraphProp = function(paraPr)
	{
		let styleId = paraPr.PStyle;
		
		if (-1 === styleId)
			displayStyleName = "";
		else if (!styleId || !styleManager.Get(styleId))
			displayStyleName = styleManager.Get(styleManager.GetDefaultParagraph()).GetName();
		else
			displayStyleName = styleManager.Get(styleId).GetName();
	}
	
	function CheckStyle(assert, expectedStyle, message)
	{
		logicDocument.UpdateInterface();
		assert.strictEqual(displayStyleName, expectedStyle.GetName(), message ? message : "");
	}

	QUnit.module("Cursor");
	
	QUnit.test("Run with/without style in paragraph", function(assert)
	{
		let word = "Word!";
		
		AscTest.ClearDocument();
		let p = AddParagraph(0, paraStyle1);
		let run = CreateRun(word, runStyle1);
		p.AddToContent(0, run);
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(false, false, 1);
		CheckStyle(assert, runStyle1, "Move cursor to the last letter of the run with a style");
		run.SetRStyle(null);
		CheckStyle(assert, paraStyle1, "Remove style from run");
	});

	QUnit.module("Select");

	QUnit.test("Add two paragraphs with different styles", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = AddParagraph(0, paraStyle1);
		assert.ok(true, "Create paragraph with paraStyle1 style");
		p1.Add_ToContent(0, CreateRun(strFirstWord));
		assert.ok(true, "Create run without style");
		let p2 = AddParagraph(1, paraStyle2);
		assert.ok(true, "Create paragraph with paraStyle2 style");
		p2.Add_ToContent(0, CreateRun(strSecondWord));
		assert.ok(true, "Create run without style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord.length);
		assert.ok(true, "Select two runs");
		CheckStyle(assert, paraStyle1, "If different paragraphs are selected, the style of the first one is visible - paraStyle1");
	});

	QUnit.test("Select part of run with style", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = AddParagraph(0, paraStyle1);
		assert.ok(true, "Create paragraph with paraStyle1 style");
		p1.Add_ToContent(0, CreateRun(strFirstWord, runStyle1));
		assert.ok(true, "Create run with runStyle1 style");
		let p2 = AddParagraph(1, paraStyle2);
		assert.ok(true, "Create paragraph with paraStyle2 style");
		p2.Add_ToContent(0, CreateRun(strSecondWord, runStyle2));
		assert.ok(true, "Create run with runStyle2 style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length  - 1);
		assert.ok(true, "Select part of first run");
		CheckStyle(assert, runStyle2, "Must show style of run - runStyle2");
	});

	QUnit.test("Select part of paragraph with style", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = AddParagraph(0, paraStyle1);
		assert.ok(true, "Create paragraph with paraStyle1 style");
		p1.Add_ToContent(0, CreateRun(strFirstWord));
		assert.ok(true, "Create run without style");
		let p2 = AddParagraph(1, paraStyle2);
		assert.ok(true, "Create paragraph with paraStyle2 style");
		p2.Add_ToContent(0, CreateRun(strSecondWord));
		assert.ok(true, "Create run without style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length  - 1);
		assert.ok(true, "Select part of Paragraph");
		CheckStyle(assert, paraStyle2, "Must show style of second selected paragraph - paraStyle2");

		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length );
		assert.ok(true, "Select all second paragraph");
		CheckStyle(assert, paraStyle2, "Must show style of second selected paragraph - paraStyle2");

		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord);
		assert.ok(true, "Select first and second paragraph");
		CheckStyle(assert, paraStyle1, "Must show style of first selected paragraph - paraStyle1");
	});

	QUnit.test("Runs with same style in one paragraph", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p = AddParagraph(0, paraStyle1);
		assert.ok(true, "Create paragraph with paraStyle1 style");
		p.Add_ToContent(0, CreateRun(strFirstWord, runStyle1));
		assert.ok(true, "Create run with runStyle1 style");
		p.Add_ToContent(1, CreateRun(strSecondWord, runStyle1));
		assert.ok(true, "Create run with runStyle1 style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length);
		assert.ok(true, "Select second Run");
		CheckStyle(assert, runStyle1, "Must show style of run - runStyle1");

		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord.length);
		assert.ok(true, "Select all runs");
		CheckStyle(assert, runStyle1, "Must show style of runs - runStyle1");
	});

	QUnit.test("Multiple runs with same style in one paragraph", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p = AddParagraph(0, paraStyle1);
		assert.ok(true, "Create paragraph with paraStyle1 style");
		p.Add_ToContent(0, CreateRun(strFirstWord, runStyle1));
		assert.ok(true, "Create run with runStyle1 style");
		p.Add_ToContent(1, CreateRun(strSecondWord, runStyle1));
		assert.ok(true, "Create another run with runStyle1 style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord.length);
		assert.ok(true, "Select all runs");
		CheckStyle(assert, runStyle1, "Must show style of runs runStyle1");
	});

	QUnit.test("Multiple runs with different style in one paragraph", function (assert)
	{
		let strFirstWord = "Word!";
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p = AddParagraph(0, paraStyle1);
		assert.ok(true, "Create paragraph with paraStyle1 style");
		p.Add_ToContent(0, CreateRun(strFirstWord, runStyle1));
		assert.ok(true, "Create run with runStyle1 style");
		p.Add_ToContent(1, CreateRun(strSecondWord, runStyle2));
		assert.ok(true, "Create run with runStyle2 style");
		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length);
		assert.ok(true, "Select second run");
		CheckStyle(assert, runStyle2, "Must show style of last run - runStyle2");

		p.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strFirstWord.length);
		assert.ok(true, "Select second run");
		CheckStyle(assert, paraStyle1, "Must show style of paragraph, so runs with different style - paraStyle1");
	});

	QUnit.test("Multiple runs with same style in two different style paragraphs", function (assert)
	{
		let strFirstWord = "Word!";
		let strDel = "Del"
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = AddParagraph(0, paraStyle1);
		assert.ok(true, "Create paragraph with paraStyle1 style");
		p1.Add_ToContent(0, CreateRun(strFirstWord, runStyle1));
		assert.ok(true, "Create run with runStyle1 style");
		p1.Add_ToContent(1, CreateRun(strDel, runStyle2));
		assert.ok(true, "Create another run with runStyle2 style");

		let p2 = AddParagraph(1, paraStyle2);
		assert.ok(true, "Create paragraph with paraStyle2 style");
		p2.Add_ToContent(0, CreateRun(strSecondWord, runStyle2));
		assert.ok(true, "Create run with runStyle2 style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strDel.length);
		assert.ok(true, "Select two last runs");
		CheckStyle(assert, runStyle2, "Must show style of two last runs - runStyle2");

		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strDel.length + strFirstWord.length);
		assert.ok(true, "Select all runs");
		CheckStyle(assert, paraStyle1, "Must show style of first Paragraph, because style of runs in selection are different - runStyle2");
	});

	QUnit.test("Space anf tab in run", function (assert)
	{
		let strFirstWord = "Word!";
		let strDel = "   	   "
		let strSecondWord = "Hello!";

		AscTest.ClearDocument();
		let p1 = AddParagraph(0, paraStyle1);
		assert.ok(true, "Create paragraph with paraStyle1 style");
		p1.Add_ToContent(0, CreateRun(strFirstWord, runStyle1));
		assert.ok(true, "Create run with runStyle1 style");
		p1.Add_ToContent(1, CreateRun(strDel, runStyle2));
		assert.ok(true, "Create another run with runStyle2 style with spaces and tabs");

		let p2 = AddParagraph(1, paraStyle2);
		assert.ok(true, "Create paragraph with paraStyle2 style");
		p2.Add_ToContent(0, CreateRun(strSecondWord, runStyle1));
		assert.ok(true, "Create run with runStyle1 style");
		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strDel.length);
		assert.ok(true, "Select two last runs");
		CheckStyle(assert, runStyle1, "Must show style of last run, runs with spaces and other style dont affect - runStyle1");

		p2.MoveCursorToEndPos();
		assert.ok(true, "Move cursor to end pos of paragraph");

		AscTest.MoveCursorLeft(true, false, strSecondWord.length + strDel.length + strFirstWord.length);
		assert.ok(true, "Select all runs");
		CheckStyle(assert, runStyle1, "Must show style of last run, runs with spaces and other style dont affect - runStyle1");
	});
});
