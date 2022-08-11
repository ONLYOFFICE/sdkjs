/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
	logicDocument.Start_SilentMode();

	logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);

	let formsManager = logicDocument.GetFormsManager();

	let p1 = new AscWord.CParagraph(editor.WordControl);
	let p2 = new AscWord.CParagraph(editor.WordControl);

	logicDocument.AddToContent(0, p1);
	logicDocument.AddToContent(1, p2);

	let r1 = new AscWord.CRun();
	p1.AddToContent(0, r1);
	r1.AddText("Hello Word!");

	let r2 = new AscWord.CRun();
	p2.AddToContent(0, r2);
	r2.AddText("Абракадабра");


	QUnit.module("Check complex forms");


	QUnit.test("Test: \"positioning, moving cursor and adding/removing text\"", function (assert)
	{
		let complexForm = logicDocument.AddComplexForm();
		complexForm.SetFormPr(new AscWord.CSdtFormPr());

		assert.strictEqual(formsManager.GetAllForms().length, 1, "Add complex form to document (check forms count)");

		r2.SetThisElementCurrent();
		r2.MoveCursorToStartPos();

		assert.strictEqual(complexForm.IsThisElementCurrent(), false, "Check cursor position in complex field");
		assert.strictEqual(r2.IsThisElementCurrent(), true, "Check cursor position in run");

		complexForm.SetThisElementCurrent();
		complexForm.MoveCursorToStartPos();

		assert.strictEqual(complexForm.IsThisElementCurrent(), true, "Check cursor position in complex field");
		assert.strictEqual(r2.IsThisElementCurrent(), false, "Check cursor position in run");

		assert.strictEqual(complexForm.IsPlaceHolder(), true, "Is placeholder in complexForm");

		// Наполняем нашу форму: 111<textForm>222<textForm>333

		let tempRun1 = new AscWord.CRun();
		tempRun1.AddText("111");
		complexForm.Add(tempRun1);
		assert.strictEqual(complexForm.IsCursorAtEnd(), true, "Check cursor after adding text");
		assert.strictEqual(complexForm.IsPlaceHolder(), false, "Check placeholder in complexForm after adding text run");

		let textForm1 = logicDocument.AddContentControlTextForm();
		textForm1.SetFormPr(new AscWord.CSdtFormPr());

		logicDocument.RemoveSelection();
		complexForm.SetThisElementCurrent();
		complexForm.MoveCursorToEndPos();

		let tempRun2 = new AscWord.CRun();
		tempRun2.AddText("222");
		complexForm.Add(tempRun2);

		let textForm2 = logicDocument.AddContentControlTextForm();
		textForm2.SetFormPr(new AscWord.CSdtFormPr());

		complexForm.SetThisElementCurrent();
		complexForm.MoveCursorToEndPos();

		let tempRun3 = new AscWord.CRun();
		tempRun3.AddText("333");
		complexForm.Add(tempRun3);

		assert.strictEqual(formsManager.GetAllForms().length, 1, "Check forms count after adding 2 subforms");

		logicDocument.RemoveSelection();
		textForm1.SetThisElementCurrent();
		assert.strictEqual(textForm1.IsThisElementCurrent(), true, "Check cursor position in textForm1");
		assert.strictEqual(complexForm.IsThisElementCurrent(), true, "Check cursor position in complex field");
		assert.strictEqual(textForm1.IsPlaceHolder(), true, "Is placeholder in textForm1");

		textForm1.Add(new AscWord.CRunText(0x61));
		textForm1.Add(new AscWord.CRunText(0x62));
		textForm1.Add(new AscWord.CRunText(0x63));
		assert.strictEqual(textForm1.IsPlaceHolder(), false, "Is placeholder in textForm1 after adding text");

		textForm2.Add(new AscWord.CRunText(0x64));
		textForm2.Add(new AscWord.CRunText(0x65));
		textForm2.Add(new AscWord.CRunText(0x66));
		assert.strictEqual(textForm2.IsPlaceHolder(), false, "Is placeholder in textForm2 after adding text");

		AscTest.SetEditingMode(logicDocument);
		assert.strictEqual(logicDocument.IsFillingFormMode(), false, "Check normal editing mode");

		textForm2.SetThisElementCurrent();
		textForm2.MoveCursorToStartPos();
		assert.strictEqual(textForm2.IsThisElementCurrent() && textForm2.IsCursorAtBegin(), true, "Move cursor at the start of textForm2");

		// Делаем два смещения, потому что после одинарного мы могли попасть в пустой ран между tempRun2 и textForm2
		// везде далее проверяем также
		logicDocument.MoveCursorLeft(false, false);
		logicDocument.MoveCursorLeft(false, false);
		assert.strictEqual(tempRun2.IsThisElementCurrent(), true, "Cursor must be in run2");

		textForm2.SetThisElementCurrent();
		textForm2.MoveCursorToEndPos();
		assert.strictEqual(textForm2.IsThisElementCurrent() && textForm2.IsCursorAtEnd(), true, "Move cursor at the end of textForm2");

		logicDocument.MoveCursorRight(false, false, false);
		logicDocument.MoveCursorRight(false, false, false);
		assert.strictEqual(tempRun3.IsThisElementCurrent(), true, "Cursor must be in run3");

		AscTest.SetFillingFormMode(logicDocument);
		assert.strictEqual(logicDocument.IsFillingFormMode(), true, "Check filling form mode");

		textForm2.SetThisElementCurrent();
		textForm2.MoveCursorToStartPos();

		logicDocument.MoveCursorLeft(false, false);
		assert.strictEqual(textForm1.IsThisElementCurrent() && textForm1.IsCursorAtEnd(), true, "Cursor must be at the end of text form1");

		textForm2.SetThisElementCurrent();
		textForm2.MoveCursorToEndPos();
		assert.strictEqual(textForm2.IsThisElementCurrent() && textForm2.IsCursorAtEnd(), true, "Move cursor at the end of textForm2");

		logicDocument.MoveCursorRight(false, false, false);
		logicDocument.MoveCursorRight(false, false, false);
		assert.strictEqual(textForm2.IsThisElementCurrent() && textForm2.IsCursorAtEnd(), true, "Cursor must be at the end of text form2");

	});
});
