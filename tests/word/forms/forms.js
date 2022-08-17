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

	function AddFormPr(oCC)
	{
		oCC.SetFormPr(new AscWord.CSdtFormPr());
	}

	QUnit.module("Check forms");

	QUnit.test("Check text form formats", function (assert)
	{
		let format = new AscWord.CTextFormFormat();
		assert.strictEqual(format.Check("qwe123"), true, "Check no format validation");

		format.SetSymbols("1234567890");
		assert.strictEqual(format.Check("qwe123"), false, "Add symbols and check 'qwe123'");
		assert.strictEqual(format.Check("555123"), true, "Add symbols and check '555123'");

		format.SetSymbols();
		format.SetDigit();
		assert.strictEqual(format.Check("qwe123"), false, "Digit: check 'qwe123'");
		assert.strictEqual(format.Check("555123"), true, "Digit: check '555123'");

		format.SetSymbols("153");
		format.SetDigit();
		assert.strictEqual(format.Check("qwe123"), false, "Digit and symbols: check 'qwe123'");
		assert.strictEqual(format.Check("555123"), false, "Digit and symbols: check '555123'");
		assert.strictEqual(format.Check("513513"), true, "Digit and symbols: check '513513'");

		format.SetSymbols();
		format.SetLetter();
		assert.strictEqual(format.Check("qwe123"), false, "Letter: check 'qwe123'");
		assert.strictEqual(format.Check("555123"), false, "Letter: check '555123'");
		assert.strictEqual(format.Check("АБВГДЕabcdef"), true, "Letter: check 'АБВГДЕabcdef'");

		format.SetSymbols("абвгдеёжзийклмнопрстуфхцчшщъыьэюя");
		format.SetLetter();
		assert.strictEqual(format.Check("АБВГДЕabcdef"), false, "Letter and symbols: check 'АБВГДЕabcdef'");
		assert.strictEqual(format.Check("привет"), true, "Letter: check 'привет'");
		assert.strictEqual(format.Check("hello"), false, "Letter: check 'hello'");

		format.SetMask("(999)-99-9999");
		format.SetSymbols();
		assert.strictEqual(format.Check("123-12-1234"), false, "Mask (999)-99-9999: check '123-12-1234'");
		assert.strictEqual(format.Check("(123)"), true, "Mask (999)-99-9999: check '(123)'");
		assert.strictEqual(format.Check("(123)abc"), false, "Mask (999)-99-9999: check '(123)abc'");
		assert.strictEqual(format.Check("(123)-12-abc5"), false, "Mask (999)-99-9999: check '(123)-12-abc5'");
		assert.strictEqual(format.Check("(123)-12-5555"), true, "Mask (999)-99-9999: check '(123)-12-5555'");

		format.SetMask("(9^99)-99-9999");

		assert.strictEqual(format.Check("(1"), true, "Mask (9^99)-99-9999: check '(1'");
		assert.strictEqual(format.Check("(123)-12-5555"), false, "Mask (9^99)-99-9999: check '(123)-12-5555'");
		assert.strictEqual(format.Check("(193)-12-5555"), true, "Mask (9^99)-99-9999: check '(193)-12-5555'");

		format.SetMask("^aabc*");
		assert.strictEqual(format.Check("aabcd"), true, "Mask ^aabc*: check 'aabcd'");
		assert.strictEqual(format.Check("qqbcd"), false, "Mask ^aabc*: check 'qqbcd'");
		assert.strictEqual(format.Check("aqbc1"), true, "Mask ^aabc*: check 'aqbc1'");
		assert.strictEqual(format.Check("aqbc123"), false, "Mask ^aabc*: check 'aqbc123'");

		format.SetRegExp("^[A-Fa-f0-9]+$");
		assert.strictEqual(format.Check("12FF"), true, "RegExp ^[A-Fa-f0-9]+$: check '12FF'");
		assert.strictEqual(format.Check("Test"), false, "RegExp ^[A-Fa-f0-9]+$: check 'Test'");
		assert.strictEqual(format.Check("FE19FF"), true, "RegExp ^[A-Fa-f0-9]+$: check 'FE19FF'");

	});

	QUnit.test("Test: \"GetAllForms\"", function (assert)
	{
		let forms = formsManager.GetAllForms();
		assert.strictEqual(forms.length, 0, "Check forms count (must be zero)");

		logicDocument.MoveCursorToStartPos();

		AddFormPr(logicDocument.AddContentControlCheckBox());
		forms = formsManager.GetAllForms();
		assert.strictEqual(forms.length, 1, "Check forms count after adding checkbox form");

		AddFormPr(logicDocument.AddContentControlComboBox());
		forms = formsManager.GetAllForms();
		assert.strictEqual(forms.length, 2, "Check forms count after adding combobox form");

		logicDocument.AddContentControlComboBox();
		forms = formsManager.GetAllForms();
		assert.strictEqual(forms.length, 2, "Check forms count after adding combobox content control");
	});
});
