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

	let strSentenceCase		= "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec tincidunt cursus lectus venenatis eleifend. Curabitur pretium pretium purus, vitae ultricies enim.";
	let strUpperCase		= "LOREM IPSUM DOLOR SIT AMET, CONSECTETUR ADIPISCING ELIT. DONEC TINCIDUNT CURSUS LECTUS VENENATIS ELEIFEND. CURABITUR PRETIUM PRETIUM PURUS, VITAE ULTRICIES ENIM.";
	let strLowerCase		= "lorem ipsum dolor sit amet, consectetur adipiscing elit. donec tincidunt cursus lectus venenatis eleifend. curabitur pretium pretium purus, vitae ultricies enim.";
	let strToggleCase		= "lOREM IPSUM DOLOR SIT AMET, CONSECTETUR ADIPISCING ELIT. dONEC TINCIDUNT CURSUS LECTUS VENENATIS ELEIFEND. cURABITUR PRETIUM PRETIUM PURUS, VITAE ULTRICIES ENIM.";
	let strCapitalizeWords	= "Lorem Ipsum Dolor Sit Amet, Consectetur Adipiscing Elit. Donec Tincidunt Cursus Lectus Venenatis Eleifend. Curabitur Pretium Pretium Purus, Vitae Ultricies Enim.";

	let logicDocument = AscTest.CreateLogicDocument();

	QUnit.module("Text");

	QUnit.test("Sentence case paragraph", function (assert)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strLowerCase);

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, strLowerCase.length);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.SentenceCase);

		assert.strictEqual(
			p.GetText().trim(),
			strSentenceCase,
			"Check sentence case"
		);
	});

	QUnit.test("Upper case paragraph", function (assert)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strSentenceCase);

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, strSentenceCase.length);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.UpperCase);

		assert.strictEqual(
			p.GetText().trim(),
			strUpperCase,
			"Check upper case"
		);
	});

	QUnit.test("Lower case paragraph", function (assert)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strToggleCase);

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, strToggleCase.length);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.LowerCase);

		assert.strictEqual(
			p.GetText().trim(),
			strLowerCase,
			"Check lower case"
		);
	});

	QUnit.test("Toggle case paragraph", function (assert)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strSentenceCase);

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, strSentenceCase.length);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.ToggleCase);

		assert.strictEqual(
			p.GetText().trim(),
			strToggleCase,
			"Check toggle case"
		);
	});

	QUnit.test("CapitalizeWords case paragraph", function (assert)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strSentenceCase);

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, strSentenceCase.length);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.CapitalizeWords);

		assert.strictEqual(
			p.GetText().trim(),
			strCapitalizeWords,
			"Check CapitalizeWords"
		);
	});

	QUnit.test("Sentence case", function (assert)
	{
		let strTest = 'two. hello'
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strTest);
		AscTest.MoveCursorToParagraph(p, true)

		AscTest.MoveCursorRight(true, false, 8);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.SentenceCase);

		assert.strictEqual(
			p.GetText().trim(),
			'Two. Hello',
			"Check sentence case"
		);
	});

	QUnit.test("Upper case", function (assert)
	{
		let strTest = 'One two. hello'
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strTest);
		AscTest.MoveCursorToParagraph(p, true)

		AscTest.MoveCursorRight(false, false, 4);
		AscTest.MoveCursorRight(true, false, 3);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.UpperCase);

		assert.strictEqual(
			p.GetText().trim(),
			'One TWO. hello',
			"Check upper case"
		);
	});

	QUnit.test("Lower case", function (assert)
	{
		let strTest = 'One Two. hello'
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strTest);
		AscTest.MoveCursorToParagraph(p, true)

		AscTest.MoveCursorRight(false, false, 4);
		AscTest.MoveCursorRight(true, false, 3);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.LowerCase);

		assert.strictEqual(
			p.GetText().trim(),
			'One two. hello',
			"Check upper case"
		);
	});

	QUnit.test("Toggle case", function (assert)
	{
		let strTest = 'One Two. hello'
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strTest);
		AscTest.MoveCursorToParagraph(p, true)

		AscTest.MoveCursorRight(false, false, 4);
		AscTest.MoveCursorRight(true, false, 3);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.ToggleCase);

		assert.strictEqual(
			p.GetText().trim(),
			'One tWO. hello',
			"Check toggle case"
		);
	});

	QUnit.test("CapitalizeWords case", function (assert)
	{
		let strTest = 'one two. hello'
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		AscTest.AddTextToParagraph(p, strTest);
		AscTest.MoveCursorToParagraph(p, true)

		AscTest.MoveCursorRight(true, false, 7);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.CapitalizeWords);

		assert.strictEqual(
			p.GetText().trim(),
			'One Two. hello',
			"Check CapitalizeWords case"
		);
	});

	QUnit.module("Math");

	function AddText(m, str)
	{
		let one = str.getUnicodeIterator();

		while (one.isInside()) {
			let oElement = new AscWord.CRunText(one.value());
			m.Add(oElement);
			one.next();
		}
	}

	QUnit.test("Sentence case math", function (assert)
	{
		let m = AscTest.CreateMath();

		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		p.AddToContent(0, m);
		AddText(m, "(abc. aaaaa)/2 ");

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, 30);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.SentenceCase);

		// in math punctuation dot t have so importance, as in regular text
		// and (3.xxx) may could be a number

		assert.strictEqual(
			m.Root.GetTextOfElement().GetText(),
			"(abc. aaaaa)/2",
			"Check sentence case"
		);
	});

	QUnit.test("Upper case math", function (assert)
	{
		let m = AscTest.CreateMath();

		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		p.AddToContent(0, m);
		AddText(m, "abc/def+2_(xyz.rt\\delta aaa+2) ");

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, 30);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.UpperCase);

		assert.strictEqual(
			m.Root.GetTextOfElement().GetText(),
			'abc/def+2_(xyz.rtδaaa+2)',
			"Check upper case"
		);
	});

	QUnit.test("Lower case math", function (assert)
	{
		let m = AscTest.CreateMath();

		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		p.AddToContent(0, m);
		AddText(m, "ABC/DEF+2_(XYZ.RTΔAAA+2) ");

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, 30);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.LowerCase);

		assert.strictEqual(
			m.Root.GetTextOfElement().GetText(),
			'ABC/DEF+2_(XYZ.RTΔAAA+2)',
			"Check upper case"
		);
	});

	QUnit.test("ToggleCase case math", function (assert)
	{
		let m = AscTest.CreateMath();

		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		p.AddToContent(0, m);
		AddText(m, "aBC/Def+2_(XyZ.RtΔAaA+2) ");

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, 30);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.ToggleCase);

		assert.strictEqual(
			m.Root.GetTextOfElement().GetText(),
			'aBC/Def+2_(XyZ.RtΔAaA+2)',
			"Check ToggleCase case"
		);
	});

	QUnit.test("CapitalizeWords case math", function (assert)
	{
		let m = AscTest.CreateMath();

		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		p.AddToContent(0, m);
		AddText(m, "aBC/Def+2_(XyZ.RtΔAaA+2) ");

		AscTest.MoveCursorToParagraph(p, true)
		AscTest.MoveCursorRight(true, false, 30);

		logicDocument.ChangeTextCase(Asc.c_oAscChangeTextCaseType.CapitalizeWords);

		assert.strictEqual(
			m.Root.GetTextOfElement().GetText(),
			'aBC/Def+2_(XyZ.RtΔAaA+2)',
			"Check CapitalizeWords case"
		);
	});
})
