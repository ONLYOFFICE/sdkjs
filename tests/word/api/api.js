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

	QUnit.module("Test api for the document editor");


	QUnit.test("Test AddText/RemoveSelection", function (assert)
	{
		AscTest.ClearDocument();

		let p = new AscWord.CParagraph(AscTest.DrawingDocument);
		logicDocument.AddToContent(0, p);

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(), "", "Check empty selection");

		logicDocument.AddTextWithPr("Hello World!");

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false, {NewLineParagraph : true}), "Hello World!\r\n", "Add text 'Hello World!'");

		let p2 = new AscWord.CParagraph(AscTest.DrawingDocument);
		logicDocument.AddToContent(1, p2);

		logicDocument.RemoveSelection();
		p2.SetThisElementCurrent();
		p2.MoveCursorToStartPos();

		logicDocument.AddTextWithPr("Second paragraph");

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false, {NewLineParagraph : true}), "Hello World!\r\nSecond paragraph\r\n", "Add text to the second paragraph");

		logicDocument.AddTextWithPr("Test");
		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false, {NewLineParagraph : true}), "Test\r\n", "Replace all with adding text 'Test'");

		function StartTest(text)
		{
			AscTest.ClearDocument();
			let p = new AscWord.CParagraph(AscTest.DrawingDocument);
			logicDocument.AddToContent(0, p);
			logicDocument.AddTextWithPr(text);
			logicDocument.MoveCursorToEndPos();
		}

		function GetText()
		{
			logicDocument.SelectAll();
			let text = logicDocument.GetSelectedText();
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
});
