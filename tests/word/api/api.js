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
	
	let logicDocument = AscTest.CreateLogicDocument();
	// Выставим стандартные настройки для параграфа
	logicDocument.GetStyles().Set_DefaultParaPr(AscWord.CParaPr.fromObject({
		Spacing : {
			After : 10 * g_dKoef_pt_to_mm,
			LineRule : linerule_Auto,
			Line : 259 / 240
		}
	}));

	QUnit.module("Test api for the document editor");


	QUnit.test("Test AddText/RemoveSelection", function (assert)
	{
		AscTest.ClearDocument();

		let p = new AscWord.Paragraph();
		logicDocument.AddToContent(0, p);

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(), "\r\n", "Check empty selection");

		logicDocument.AddTextWithPr("Hello World!");

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false), "Hello World!\r\n", "Add text 'Hello World!'");

		let p2 = new AscWord.Paragraph();
		logicDocument.AddToContent(1, p2);

		logicDocument.RemoveSelection();
		p2.SetThisElementCurrent();
		p2.MoveCursorToStartPos();

		logicDocument.AddTextWithPr("Second paragraph");

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false), "Hello World!\r\nSecond paragraph\r\n", "Add text to the second paragraph");

		logicDocument.AddTextWithPr("Test");
		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false), "Test\r\n", "Replace all with adding text 'Test'");

		function StartTest(text)
		{
			AscTest.ClearDocument();
			let p = new AscWord.Paragraph();
			logicDocument.AddToContent(0, p);
			logicDocument.AddTextWithPr(text);
			logicDocument.MoveCursorToEndPos();
		}

		function GetText()
		{
			logicDocument.SelectAll();
			let text = logicDocument.GetSelectedText(false, {ParaSeparator : ""});
			logicDocument.RemoveSelection();
			return text;
		}

		let settings = new AscCommon.CAddTextSettings();

		settings.SetWrapWithSpaces(false);

		StartTest("Text");
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex123t", "Check text Tex123t");

		StartTest("Tex t");
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex 123t", "Check text Tex 123t");

		StartTest("Tex t");
		AscTest.MoveCursorLeft();
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex123 t", "Check text Tex123 t");

		StartTest("Text");
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Text123", "Check text Text123");

		StartTest("Text");
		logicDocument.MoveCursorToStartPos();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "123Text", "Check text 123Text");

		settings.SetWrapWithSpaces(true);

		StartTest("Text");
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex 123 t", "Check wrap spaces Tex 123 t");

		StartTest("Tex t");
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex 123 t", "Check wrap spaces Tex 123 t");

		StartTest("Tex t");
		AscTest.MoveCursorLeft();
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex 123 t", "Check wrap spaces Tex 123 t");

		StartTest("Text");
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Text 123", "Check wrap spaces Text 123");

		StartTest("Text");
		logicDocument.MoveCursorToStartPos();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "123 Text", "Check wrap spaces 123 Text");


	});
	
	QUnit.test("Change numbering level", function(assert)
	{
		AscTest.ClearDocument();
		
		let p1 = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p1);
		
		let p2 = AscTest.CreateParagraph();
		logicDocument.AddToContent(1, p2);
		
		AscTest.MoveCursorToParagraph(p1, true);
		AscTest.AddNumbering(1, 1);
		AscTest.MoveCursorToParagraph(p2, true);
		AscTest.AddNumbering(1, 1);
		
		let numPr1 = p1.GetNumPr();
		let numPr2 = p2.GetNumPr();
		
		assert.strictEqual(!!numPr1, true, "Check paragraph 1 numbering");
		assert.strictEqual(!!numPr2, true, "Check paragraph 2 numbering");
		assert.deepEqual(numPr1, numPr2, "Check paragraphs have the same paragraph 2 numbering");
		assert.strictEqual(numPr2.Lvl, 0, "Check numbering lvl of paragraph 2");
		
		// При двойном Enter на нулевом уровне в пустом параграфе нумерация должна убираться
		assert.strictEqual(p2.IsEmpty(), true, "Check if paragraph 2 is empty");
		
		AscTest.MoveCursorToParagraph(p2, true);
		AscTest.PressKey(AscTest.Key.enter);
		
		assert.strictEqual(!!p2.GetNumPr(), false, "Check of removing numbering from paragraph 2");
		
		
		AscTest.MoveCursorToParagraph(p2, true);
		AscTest.AddNumbering(1, 1);
		assert.strictEqual(!!p2.GetNumPr(), true, "Check paragraph 2 numbering");
		assert.deepEqual(p1.GetNumPr(), p2.GetNumPr(), "Check paragraphs have the same paragraph 2 numbering");
		
		AscTest.SetParagraphNumberingLvl(p2, 1);
		assert.deepEqual(p2.GetNumPr().Lvl, 1, "Check level of second paragraph");
		
		// При двойном Enter на ненулевом уровне в пустом параграфе уровень нумераации должен уменьшаться
		AscTest.MoveCursorToParagraph(p2, true);
		AscTest.PressKey(AscTest.Key.enter);
		
		assert.strictEqual(!!p2.GetNumPr(), true, "Check paragraph 2 numbering");
		assert.deepEqual(p2.GetNumPr().Lvl, 0, "Check level of second paragraph");
	});
	
	QUnit.test("Test add/remove space before/after paragraph", function(assert)
	{
		const epsilon = 0.001;
		
		AscTest.ClearDocument();
		let p = [];
		
		for (let i = 0; i < 3; ++i)
		{
			p[i] = AscTest.CreateParagraph();
			logicDocument.AddToContent(i, p[i]);
		}
		
		function checkSpacing(spacings)
		{
			for (let i = 0; i < 3; ++i)
			{
				let paraPr = p[i].GetCompiledParaPr(false);
				
				assert.close(paraPr.Spacing.Before, spacings[i][0], epsilon, "Check paragraph " + (i + 1) + " spacing before");
				assert.close(paraPr.Spacing.After, spacings[i][1], epsilon, "Check paragraph " + (i + 1) + " spacing after");
			}
		}
		function checkHaveSpace(haveSpaceBefore, haveSpaceAfter)
		{
			assert.strictEqual(AscTest.Editor.asc_haveSpaceBeforeParagraph(), haveSpaceBefore, "Check have space before");
			assert.strictEqual(AscTest.Editor.asc_haveSpaceAfterParagraph(), haveSpaceAfter, "Check have space after");
		}
		
		checkSpacing([
			[0, 10 * g_dKoef_pt_to_mm],
			[0, 10 * g_dKoef_pt_to_mm],
			[0, 10 * g_dKoef_pt_to_mm]
		]);
		
		AscTest.MoveCursorToParagraph(p[0]);
		checkHaveSpace(false, true);
		
		AscTest.Editor.asc_addSpaceBeforeParagraph();
		
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[0, 10 * g_dKoef_pt_to_mm],
			[0, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(true, true);
		
		AscTest.MoveCursorToParagraph(p[1]);
		AscTest.Editor.asc_removeSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[0, 0],
			[0, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(false, false);
		
		AscTest.SelectDocumentRange(0, 2);
		AscTest.Editor.asc_removeSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 0],
			[0, 0],
			[0, 0]
		]);
		checkHaveSpace(true, false);
		
		AscTest.Editor.asc_removeSpaceBeforeParagraph();
		checkSpacing([
			[0, 0],
			[0, 0],
			[0, 0]
		]);
		checkHaveSpace(false, false);
		
		AscTest.Editor.asc_addSpaceBeforeParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 0],
			[12 * g_dKoef_pt_to_mm, 0],
			[12 * g_dKoef_pt_to_mm, 0]
		]);
		checkHaveSpace(true, false);
		
		AscTest.Editor.asc_addSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(true, true);
		
		let paraStyle = AscTest.CreateParagraphStyle("SpacingStyle");
		paraStyle.SetParaPr(AscWord.CParaPr.fromObject({
			Spacing : {
				Before : 25 * g_dKoef_pt_to_mm,
				After  : 0
			}
		}));
		p[1].SetPStyle(paraStyle.GetId());
		
		AscTest.SelectDocumentRange(1, 2);
		AscTest.Editor.asc_removeSpaceBeforeParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[0, 0],
			[0, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(false, false);
		
		AscTest.Editor.asc_addSpaceBeforeParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[25 * g_dKoef_pt_to_mm, 0],
			[25 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(true, false);
		
		AscTest.Editor.asc_addSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[25 * g_dKoef_pt_to_mm, 12 * g_dKoef_pt_to_mm],
			[25 * g_dKoef_pt_to_mm, 12 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(true, true);
		
		AscTest.Editor.asc_removeSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[25 * g_dKoef_pt_to_mm, 0],
			[25 * g_dKoef_pt_to_mm, 0]
		]);
		checkHaveSpace(true, false);
		
		
	});
	
	QUnit.test("Get text/selected text", function(assert)
	{
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		
		let run = AscTest.CreateRun();
		logicDocument.AddToContent(0, p);
		p.AddToContentToEnd(run);
		run.AddText("The quick brown fox jumps over the lazy dog");
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 4);
		assert.strictEqual(logicDocument.GetSelectedText(), "", "Check no selection");
		AscTest.MoveCursorRight(true, false, 5);
		assert.strictEqual(logicDocument.GetSelectedText(), "quick", "Select word 'quick'");
		
		// TODO: Переделать заполнение формулы по-нормальному
		let math = AscTest.CreateMath();
		p.AddToContentToEnd(math);
		math.SetThisElementCurrent();
		let text = "abcd";
		for (let i = 0; i < text.length; ++i)
		{
			math.Add(new AscWord.CRunText(text.charCodeAt(i)));
		}
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(false, false, 4); // 1 сдвиг для захода в формулу
		AscTest.MoveCursorRight(true, false, 2);
		assert.strictEqual(logicDocument.GetSelectedText(), "bc", "Add math with text 'abcd' and partially select the text");
		
	});
});
