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
	logicDocument.SetBoldText = function ()
	{
		let oTextPr = logicDocument.GetCalculatedTextPr();
		if (oTextPr)
		{
			if (!logicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
			{
				logicDocument.StartAction(AscDFH.historydescription_Document_SetTextBoldHotKey);
				logicDocument.AddToParagraph(new ParaTextPr({Bold : oTextPr.Bold !== true}));
				logicDocument.UpdateInterface();
				logicDocument.FinalizeAction();
			}
		}
	}

	let versionHistory = logicDocument.Api.VersionHistory = new Asc.asc_CVersionHistory();
	versionHistory.asc_SetUserId(0);
	versionHistory.asc_SetUserName("AnonymousUser");
	versionHistory.asc_SetDateOfRevision(new Date().getTime());

	let oCurDelRecover = null;
	let oCollaborativeHistory = null;
	let arr = [];
	let arrPoints = [];
	function AddParagraph(pos)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(pos, p);
		return p;
	}
	function CreateRun(text)
	{
		let r = AscTest.CreateRun();
		r.AddText(text);
		return r;
	}
	function DelLast(nCount, isShift)
	{
		if (isShift)
		{
			AscTest.MoveCursorLeft(true, false, nCount);
			AscTest.PressKey(AscTest.Key.backspace);
		}
		else
		{
			for (let i = 0; i < nCount; i++)
				AscTest.PressKey(AscTest.Key.backspace);
		}

	}

	function Init(nPos, isMove)
	{
		UpdateChanges(nPos, isMove);
		logicDocument.CollaborativeEditing.CoHistory.InitTextRecover();

		oCollaborativeHistory	= logicDocument.CollaborativeEditing.CoHistory;
		oCurDelRecover			= AscCommon.CollaborativeEditing.CoHistory.textRecovery;
	}

	function UpdateChanges (nPoint, isMove)
	{
		let arr = [];
		let nLengthOfPoints = AscCommon.History.Points.length - 1;

		if (nPoint !== undefined)
		{
			for (let i = nLengthOfPoints; i >= nPoint; i--)
			{
				if (isMove)
					arr.push(new AscCommon.CChangesTableIdDescription());
				AscCommon.History.GetChangesFromPoint(i, arr);
			}

			arr = arr.reverse();
		}
		else
		{
			for (let i = 0; i <= nLengthOfPoints; i++)
			{
				if (isMove)
					arr.push(new AscCommon.CChangesTableIdDescription());
				AscCommon.History.GetChangesFromPoint(i, arr);
			}
		}

		AscCommon.CollaborativeEditing.CoHistory.Changes = arr;
	}

	function ShowDelText()
	{
		oCurDelRecover.RecoverDeletedText();
	}

	function UndoDelText()
	{
		oCurDelRecover.UndoRecoveredText();
	}

	function MoveToPoint(nPos)
	{
		logicDocument.CollaborativeEditing.CoHistory.NavigationRevisionHistoryByStep(nPos);
	}

	function Prev()
	{
		logicDocument.CollaborativeEditing.CoHistory.NavigationRevisionHistoryByStep(logicDocument.CollaborativeEditing.CoHistory.GetGlobalPointIndex() - 1);
	}

	function Next()
	{
		logicDocument.CollaborativeEditing.CoHistory.NavigationRevisionHistoryByStep(logicDocument.CollaborativeEditing.CoHistory.GetGlobalPointIndex() + 1);
	}

	function CheckRuns(assert, paragraph, arr)
	{
		for (let i = 0; i < paragraph.Content.length - 1; i++)
		{
			if (!arr[i])
				return;

			let name = arr[i][1] === reviewtype_Common ? "reviewtype_Common" : "reviewtype_Remove";
			let str ="\"" + arr[i][0] + "\"" + " is " + name
			let oCurrentRun = paragraph.Content[i];
				assert.strictEqual(oCurrentRun.GetReviewType(), arr[i][1], str);
		}
	}

	QUnit.testStart(function (){
		AscCommon.History.Clear();
		AscTest.ClearDocument();
		AscCommon.CollaborativeEditing.CoHistory.textRecovery = null;
		AscCommon.CollaborativeEditing.Clear();
		arrPoints = [];
	})

	QUnit.module("Unit-tests for recover deleted text");

	QUnit.test("Delete one letter", function (assert)
	{
		let strStartText = "abc";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with 'abc' text.");

		DelLast(1);
		assert.ok(true, "Delete one last letter");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "ab", "Text in run is 'ab'");

		Init();
		ShowDelText()
		assert.ok(true, "Recover deleted text");
		let recoverRun = p.Content[1];
		assert.strictEqual(recoverRun.GetReviewType(), reviewtype_Remove, "New ParaRun ReviewType is delete");
		assert.strictEqual(run.GetReviewType(), reviewtype_Common, "Old ParaRun ReviewType is common");

		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "abc", "Text in run is 'abc'");

		CheckRuns(assert, p, [
			["ab", reviewtype_Common],
			["c", reviewtype_Remove],
		]);
	});

	QUnit.test("Delete letter block (with selection)", function (assert)
	{
		let strStartText = "Hello World";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with '" + strStartText+"' text.");

		DelLast(6, true);
		assert.ok(true, "Delete ' World'");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello", "Text in run is 'Hello'");

		Init();
		ShowDelText()
		assert.ok(true, "Recover deleted text");
		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "Hello World", "Text in run is 'Hello World'");

		CheckRuns(assert, p, [
			["Hello", reviewtype_Common],
			[" World", reviewtype_Remove],
		]);
	});

	QUnit.test("Delete many letter as one block", function (assert)
	{
		let strStartText = "Hello World";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with '" + strStartText+"' text.");

		AscTest.MoveCursorLeft(false, false, 6);
		DelLast(4);
		assert.ok(true, "Delete 'llo'");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "H World", "Text in run is 'H World'");

		Init();
		ShowDelText()
		assert.ok(true, "Recover deleted text");
		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "Hello World", "Text in run is 'Hello World'");

		CheckRuns(assert, p, [
			["He", reviewtype_Common],
			["llo", reviewtype_Remove],
			[" World", reviewtype_Common],
		]);
	});

	QUnit.test("Delete letter blocks (from left to right)", function (assert)
	{
		let strStartText = "Hello World";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with '" + strStartText+"' text.");

		AscTest.MoveCursorLeft(false, false, 8);
		DelLast(2);
		assert.ok(true, "Delete 'el'");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hlo World", "Text in run is 'Hlo World'");

		AscTest.MoveCursorRight(false, false, 6);
		DelLast(2);
		assert.ok(true, "Delete 'or'");
		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hlo Wld", "Text in run is 'Hlo Wld'");

		Init();
		ShowDelText()
		assert.ok(true, "Recover deleted text");
		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "Hello World", "Text in run is 'Hello World'");

		CheckRuns(assert, p, [
			["H", reviewtype_Common],
			["el", reviewtype_Remove],
			["lo W", reviewtype_Common],
			["or", reviewtype_Remove],
			["ld", reviewtype_Common],
		]);
	});

	QUnit.test("Delete letter blocks (from right to left)", function (assert)
	{
		let strStartText = "Hello World";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with '" + strStartText+"' text.");
		AscTest.MoveCursorToParagraph(p, false);

		AscTest.MoveCursorLeft(false, false, 2);
		AscTest.MoveCursorLeft(true, false, 2);
		DelLast(1);
		assert.ok(true, "Delete 'or'");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello Wld", "Text in run is 'Hello Wld'");

		AscTest.MoveCursorLeft(false, false, 4);
		AscTest.MoveCursorLeft(true, false, 2);
		DelLast(1);
		assert.ok(true, "Delete 'el'");
		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hlo Wld", "Text in run is 'Hlo Wld'");

		Init();
		ShowDelText()
		assert.ok(true, "Recover deleted text");
		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "Hello World", "Text in run is 'Hello World'");

		CheckRuns(assert, p, [
			["H", reviewtype_Common],
			["el", reviewtype_Remove],
			["lo W", reviewtype_Common],
			["or", reviewtype_Remove],
			["ld", reviewtype_Common],
		]);
	});

	QUnit.test("Delete paragraph", function (assert)
	{
		let strOne = "One";
		let strTwo = "Two";
		let p = AddParagraph(0);
		let p2 = AddParagraph(1);
		let run = CreateRun(strOne);
		let run2 = CreateRun(strTwo);
		p.AddToContentToEnd(run);
		p2.AddToContentToEnd(run2);

		AscTest.MoveCursorToParagraph(p2, false);

		DelLast(4);
		assert.ok(true, "Delete 'Two'");

		let strDeletedText = AscTest.GetParagraphText(p2);
		assert.strictEqual(strDeletedText, "", "Text in document is ''");

		Init();
		ShowDelText()
		assert.ok(true, "Recover deleted text");
		let strResultText = AscTest.GetParagraphText(p2);
		assert.strictEqual(strResultText, "Two", "Text in run is 'One'");

		CheckRuns(assert, p, [
			["One", reviewtype_Common],
		]);

		CheckRuns(assert, p2, [
			["Two", reviewtype_Remove],
		]);
	});



	QUnit.test("Going back and forth through history - one letter", function (assert)
	{
		let strStartText = "abc";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with 'abc' text.");

		DelLast(1);
		assert.ok(true, "Delete one last letter");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "ab", "Text in run is 'ab'");

		Init();
		Prev();
		assert.ok(true, "Prev in history");

		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "abc", "Text in run is 'abc'");

		Next();
		assert.ok(true, "Next in history");

		let strResultText2 = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText2, "ab", "Text in run is 'ab'");
	});

	QUnit.test("Going back and forth through history - letter block", function (assert)
	{
		let strStartText = "Hello World";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with '" + strStartText+"' text.");

		DelLast(6);
		assert.ok(true, "Delete ' World'");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello", "Text in run is 'Hello'");

		Init();
		Prev();
		assert.ok(true, "Prev in history");

		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "Hello World", "Text in run is 'Hello World'");

		Next();
		assert.ok(true, "Next in history");

		let strResultText2 = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText2, "Hello", "Text in run is 'Hello'");
	});

	QUnit.test("Going back and forth through history - letter blocks (deleted from left to right)", function (assert)
	{
		let strStartText = "Hello World";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with '" + strStartText+"' text.");

		AscTest.MoveCursorLeft(false, false, 8);
		DelLast(2);
		assert.ok(true, "Delete 'el'");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hlo World", "Text in run is 'Hlo World'");

		AscTest.MoveCursorRight(false, false, 6);
		DelLast(2);
		assert.ok(true, "Delete 'or'");
		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hlo Wld", "Text in run is 'Hlo Wld'");

		Init();
		Prev();
		assert.ok(true, "Prev in history");

		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "Hello World", "Text in run is 'Hello World'");

		Next();
		assert.ok(true, "Next in history");

		let strResultText2 = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText2, "Hlo Wld", "Text in run is 'Hlo Wld'");

	});

	QUnit.test("Going back and forth through history - letter blocks (from right to left)", function (assert)
	{
		let strStartText = "Hello World";
		let p = AddParagraph(0);
		let run = CreateRun(strStartText);
		p.AddToContentToEnd(run);
		assert.ok(true, "Create run with '" + strStartText+"' text.");
		AscTest.MoveCursorToParagraph(p, false);

		AscTest.MoveCursorLeft(false, false, 2);
		AscTest.MoveCursorLeft(true, false, 2);
		DelLast(1);
		assert.ok(true, "Delete 'or'");
		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello Wld", "Text in run is 'Hello Wld'");

		AscTest.MoveCursorLeft(false, false, 4);
		AscTest.MoveCursorLeft(true, false, 2);
		DelLast(1);
		assert.ok(true, "Delete 'el'");
		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hlo Wld", "Text in run is 'Hlo Wld'");

		Init();
		Prev();
		assert.ok(true, "Prev in history");

		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "Hello World", "Text in run is 'Hello World'");

		Next();
		assert.ok(true, "Next in history");

		let strResultText2 = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText2, "Hlo Wld", "Text in run is 'Hlo Wld'");
	});

	QUnit.test("Going back and forth through history - paragraph", function (assert)
	{
		let strOne = "One";
		let p = AddParagraph(0);
		let run = CreateRun(strOne);
		p.AddToContentToEnd(run);

		DelLast(5);
		assert.ok(true, "Delete 'One'");

		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "", "Text in document is ''");

		Init();
		Prev();
		assert.ok(true, "Prev in history");

		let strResultText = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText, "One", "Text in run is 'One'");

		Next();
		assert.ok(true, "Next in history");

		let strResultText2 = AscTest.GetParagraphText(p);
		assert.strictEqual(strResultText2, "", "Text in run is ''");
	});

	QUnit.test("Complex 1", function (assert)
	{
		logicDocument.AddToContent(0, AscTest.CreateParagraph());

		AscTest.EnterText("Hello hello");
		assert.ok(true, "Input 'Hello hello' in first paragraph");

		AscTest.PressKey(13);
		assert.ok(true, "Press Enter for create second paragraph");

		AscTest.EnterText("Hello world 123");
		assert.ok(true, "Input 'Hello world 123' in second paragraph");

		AscTest.MoveCursorLeft(false, false, 3);
		assert.ok(true, "Move cursor to start pos of '123'");

		AscTest.MoveCursorLeft(true, false, 5);
		assert.ok(true, "Select 'world'");

		DelLast(1, true);
		assert.ok(true, "Delete 'world'");

		let p = logicDocument.Content[0];
		let p2 = logicDocument.Content[1];

		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello hello", "Text in first paragraph 'Hello hello'");

		let strDeletedText2 = AscTest.GetParagraphText(p2);
		assert.strictEqual(strDeletedText2, "Hello 123", "Text in second paragraph 'Hello 123'");

		Init(3);

		ShowDelText();
		assert.ok(true, "Show del text");

		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello hello", "Text in run is 'Hello hello'");

		strDeletedText2 = AscTest.GetParagraphText(p2);
		assert.strictEqual(strDeletedText2, "Hello world 123", "Text in run is 'Hello world 123'");

		CheckRuns(assert, p2, [
			["Hello ", reviewtype_Common],
			["world ", reviewtype_Remove],
			["123", reviewtype_Common],
		]);

		UndoDelText();
		assert.ok(true, "Undo show del text");

		CheckRuns(assert, p2, [
			["Hello 123", reviewtype_Common],
		]);
	});

	QUnit.test("Complex 2", function (assert)
	{
		logicDocument.AddToContent(0, AscTest.CreateParagraph());

		AscTest.EnterText("Hello hello");
		assert.ok(true, "Input 'Hello hello' in first paragraph");

		AscTest.MoveCursorLeft(false, false, 2);
		DelLast(3, false);
		assert.ok(true, "Delete from first paragraph 'hel'");

		let p = logicDocument.Content[0];

		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello lo", "Text in first paragraph 'Hello hello'");

		Init(1);
		ShowDelText();

		assert.ok(true, "Show del text");

		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello hello", "Text in run is 'Hello hello'");

		CheckRuns(assert, p, [
			["Hello ", reviewtype_Common],
			["hel", reviewtype_Remove],
			["lo", reviewtype_Common],
		]);

		UndoDelText();
		assert.ok(true, "Undo show del text");

		CheckRuns(assert, p, [
			["Hello lo", reviewtype_Common],
		]);
	});
	QUnit.test("Complex 3", function (assert)
	{
		logicDocument.AddToContent(0, AscTest.CreateParagraph());

		AscTest.EnterText("Hello");
		AscTest.EnterText(" hello");

		assert.ok(true, "Input 'Hello hello' in first paragraph");

		let p = logicDocument.Content[0];

		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello hello", "Text in first paragraph 'Hello hello'");

		Init(undefined, true);
		Prev();
		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello", "Text in run is 'Hello'");

		CheckRuns(assert, p, [
			["Hello ", reviewtype_Common],
		]);
	});
	QUnit.test("Complex 4", function (assert)
	{
		logicDocument.AddToContent(0, AscTest.CreateParagraph());

		AscTest.EnterText("Hello3");
		DelLast(1, false);
		AscTest.EnterText(" word");
		DelLast(5, false);
		AscTest.EnterText(" world");

		assert.ok(true, "* 𝐛𝐨𝐥𝐝 𝐭𝐞𝐱𝐭 𝐢𝐬 𝐝𝐞𝐥𝐞𝐭𝐞𝐝");
		assert.ok(true, "Input 'Hello𝟑 𝐰𝐨𝐫𝐝 world'");

		let p = logicDocument.Content[0];

		let strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello world", "Input 'Hello𝟑 𝐰𝐨𝐫𝐝 world'");

		Init(undefined, true);

		MoveToPoint(8)
		assert.ok(true, "Move to point 9");

		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello", "Input 'Hello'");

		MoveToPoint(3)
		assert.ok(true, "Move to point 4");

		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello word", "Input 'Hello word'");

		MoveToPoint(1)
		assert.ok(true, "Move to point 2");

		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello3", "Input 'Hello3'");

		MoveToPoint(8)
		assert.ok(true, "Move to point 9");

		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello", "Input 'Hello'");

		MoveToPoint(3)
		assert.ok(true, "Move to point 4");

		strDeletedText = AscTest.GetParagraphText(p);
		assert.strictEqual(strDeletedText, "Hello word", "Input 'Hello word'");
	});

	QUnit.test("Split run", function (assert)
	{
		logicDocument.AddToContent(0, AscTest.CreateParagraph());

		AscTest.EnterText("Hello how are you");
		let p = logicDocument.Content[0];
		AscTest.SelectParagraphRange(p,3, 11);

		logicDocument.SetBoldText();

		let r0 = logicDocument.Content[0].Content[0];
		let r1 = logicDocument.Content[0].Content[1];
		let r2 = logicDocument.Content[0].Content[2];

		assert.strictEqual(r0.Pr.Bold, undefined, "Check \"Hel\" not bold");
		assert.strictEqual(r1.Pr.Bold, true, "Check \"lo how a\" is bold");
		assert.strictEqual(r2.Pr.Bold, undefined, "Check \"re you\" not bold");

		AscTest.SelectParagraphRange(p,5, 9);
		DelLast(1, false);

		assert.strictEqual(AscTest.GetParagraphText(p), "Hello are you", "Text is \"Hello are you\"");

		Init(2);
		ShowDelText();
		assert.ok(true, 'Show del text');

		CheckRuns(assert, p, [
			["Hel",		reviewtype_Common],
			["lo",		reviewtype_Common],
			[" how",	reviewtype_Remove],
			[" a",		reviewtype_Common],
			["re you",	reviewtype_Common],
		]);
	});

	QUnit.test("Check is not show del text in context of one revision", function (assert)
	{
		logicDocument.AddToContent(0, AscTest.CreateParagraph());

		AscTest.EnterText("Hello");
		DelLast(5, false);
		AscTest.EnterText("World");
		let p = logicDocument.Content[0];

		assert.strictEqual(AscTest.GetParagraphText(p), "World", "Text is \"World\"");

		Init(0);
		ShowDelText();
		assert.ok(true, 'Show del text');
		assert.strictEqual(AscTest.GetParagraphText(p), "World", "Text is \"World\", 'Hello' not show");

		CheckRuns(assert, p, [
			["World",		reviewtype_Common],
		]);
	});
});
