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

"use strict";

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();
	const numberingManager = logicDocument.GetNumberingManager();
	logicDocument.Start_SilentMode();

	function GenerateClearParagraph(enterText)
	{
		AscTest.ClearDocument();
		const p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.EnterText(enterText);
		return p;
	}

	QUnit.module("Check autocorrect text to numbering in clear paragraph");
	QUnit.test("Check autocorrect text to numbering in clear paragraph", (assert) =>
	{
		function TestAutoCorrect(enterText, checkNumberingText, description)
		{
			const p = GenerateClearParagraph(enterText);
			assert.true(p.Is_Empty(), "Check empty paragraph after autocorrection");
			assert.strictEqual(p.GetNumberingText(false), checkNumberingText, description);
		}
		TestAutoCorrect("* ", String.fromCharCode(0x00B7), "Check bullet autocorrect with *");
		TestAutoCorrect("- ", String.fromCharCode(0x2013), "Check bullet autocorrect with -");
		TestAutoCorrect("> ", String.fromCharCode(0x00D8), "Check bullet autocorrect with >");

		TestAutoCorrect("1. ", "1.", "Check decimal with dot autocorrect for 1 level");
		TestAutoCorrect("1.1. ", "1.1.", "Check decimal with dot autocorrect for 2 levels");
		TestAutoCorrect("1.1.1. ", "1.1.1.", "Check decimal with dot autocorrect for 3 levels");
		TestAutoCorrect("1.1.1.1. ", "1.1.1.1.", "Check decimal with dot autocorrect for 4 levels");
		TestAutoCorrect("1.1.1.1.1. ", "1.1.1.1.1.", "Check decimal with dot autocorrect for 5 levels");
		TestAutoCorrect("1.1.1.1.1.1. ", "1.1.1.1.1.1.", "Check decimal with dot autocorrect for 6 levels");
		TestAutoCorrect("1.1.1.1.1.1.1. ", "1.1.1.1.1.1.1.", "Check decimal with dot autocorrect for 7 levels");
		TestAutoCorrect("1.1.1.1.1.1.1.1. ", "1.1.1.1.1.1.1.1.", "Check decimal with dot autocorrect for 8 levels");
		TestAutoCorrect("1.1.1.1.1.1.1.1.1. ", "1.1.1.1.1.1.1.1.1.", "Check decimal with dot autocorrect for 9 levels");

		TestAutoCorrect("1) ", "1)", "Check decimal with bracket autocorrect for 1 level");
		TestAutoCorrect("1)1) ", "1)1)", "Check decimal with bracket autocorrect for 2 levels");
		TestAutoCorrect("1)1)1) ", "1)1)1)", "Check decimal with bracket autocorrect for 3 levels");
		TestAutoCorrect("1)1)1)1) ", "1)1)1)1)", "Check decimal with bracket autocorrect for 4 levels");
		TestAutoCorrect("1)1)1)1)1) ", "1)1)1)1)1)", "Check decimal with bracket autocorrect for 5 levels");
		TestAutoCorrect("1)1)1)1)1)1) ", "1)1)1)1)1)1)", "Check decimal with bracket autocorrect for 6 levels");
		TestAutoCorrect("1)1)1)1)1)1)1) ", "1)1)1)1)1)1)1)", "Check decimal with bracket autocorrect for 7 levels");
		TestAutoCorrect("1)1)1)1)1)1)1)1) ", "1)1)1)1)1)1)1)1)", "Check decimal with bracket autocorrect for 8 levels");
		TestAutoCorrect("1)1)1)1)1)1)1)1)1) ", "1)1)1)1)1)1)1)1)1)", "Check decimal with bracket autocorrect for 9 levels");
		TestAutoCorrect("a. ", "a.", "Check lower letter with dot autocorrect");
		TestAutoCorrect("a) ", "a)", "Check lower letter with bracket autocorrect");
		TestAutoCorrect("A. ", "A.", "Check upper letter with dot autocorrect");
		TestAutoCorrect("A) ", "A)", "Check upper letter with bracket autocorrect");

		// todo
		// TestAutoCorrect("I. ", "I.", "Check upper roman with dot");
		// TestAutoCorrect("I) ", "I)", "Check upper roman with bracket");
		// TestAutoCorrect("i. ", "i.", "Check lower roman with dot");
		// TestAutoCorrect("i) ", "i)", "Check lower roman with bracket");
	});

	QUnit.test("Check not autocorrect text to numbering in clear paragraph", (assert) =>
	{
		function TestNotAutoCorrect(enterText, checkText, description)
		{
			const p = GenerateClearParagraph(enterText);
			assert.true(!p.GetNumPr(), "Check empty num pr");
			assert.strictEqual(AscTest.GetParagraphText(p), checkText, description);
		}

		TestNotAutoCorrect("a.a. ", "a.a. ", "Check lower letter with dot");
		TestNotAutoCorrect("a)a) ", "a)a) ", "Check lower letter with bracket");
		TestNotAutoCorrect("A.A. ", "A.A. ", "Check upper letter with dot");
		TestNotAutoCorrect("A)A) ", "A)A) ", "Check upper letter with bracket");
		TestNotAutoCorrect("I.I. ", "I.I. ", "Check upper roman with dot");
		TestNotAutoCorrect("I)I) ", "I)I) ", "Check upper roman with bracket");
		TestNotAutoCorrect("i.i. ", "i.i. ", "Check lower roman with dot");
		TestNotAutoCorrect("i)i) ", "i)i) ", "Check lower roman with bracket");
		TestNotAutoCorrect("1)1)1)1)1)1)1)1)1)1) ", "1)1)1)1)1)1)1)1)1)1) ", "Check decimal with 10 levels");
		TestNotAutoCorrect("1.1.1.1.1.1.1.1.1.1. ", "1.1.1.1.1.1.1.1.1.1. ", "Check decimal with 10 levels");
		TestNotAutoCorrect("1 ", "1 ", "Check just number");
	});

	QUnit.module("Check autocorrect text to numbering with previous numbering paragraphs");
	QUnit.test("Check autocorrect text to numbering", (assert) =>
	{
		function CreateNum(subtype)
		{
			let numInfo = AscWord.GetNumberingObjectByDeprecatedTypes(1, subtype);
			let num = numberingManager.CreateNum();
			numInfo.FillNum(num);
			numberingManager.AddNum(num);
			return num;
		}

		function GenerateDocumentWithNumberingParagraphs(subtype)
		{
			AscTest.ClearDocument();
			let num = CreateNum(subtype);
			let numPr = new AscWord.NumPr(num.GetId(), 0);
			for (let i = 0; i < 4; i += 1)
			{
				const p = AscTest.CreateParagraph();
				logicDocument.PushToContent(p);
				p.SetNumPr(numPr.NumId, numPr.Lvl);
			}
		}

		function TestAutoCorrect(subType, enterText, checkNumberText, description)
		{
			GenerateDocumentWithNumberingParagraphs(subType);
			const p = AscTest.CreateParagraph();
			logicDocument.PushToContent(p);
			AscTest.MoveCursorToParagraph(p, true);
			AscTest.EnterText(enterText);
			assert.true(p.Is_Empty(), "Check empty paragraph after autocorrection");
			assert.strictEqual(p.GetNumberingText(false), checkNumberText, description);
		}

		TestAutoCorrect(1, "5. ", "5.", "Check decimal with dot autocorrection");
		TestAutoCorrect(2, "5) ", "5)", "Check decimal with bracket autocorrection");
		TestAutoCorrect(4, "E. ", "E.", "Check upper letter with dot autocorrection");
		TestAutoCorrect(5, "e) ", "e)", "Check lower letter with bracket autocorrection");
		TestAutoCorrect(6, "e. ", "e.", "Check lower letter with dot autocorrection");

		// todo
		// TestAutoCorrect(3, "V. ", "V.", "Check upper roman with dot autocorrection");
		// TestAutoCorrect(7, "v. ", "v.", "Check lower roman with dot autocorrection");
	});
});
