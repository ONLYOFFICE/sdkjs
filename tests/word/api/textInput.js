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
	
	function addShape(x, y, h, w)
	{
		const drawing = new ParaDrawing(w, h, null, logicDocument.GetDrawingDocument(), logicDocument, null);
		const shapeTrack = new AscFormat.NewShapeTrack('rect', x, y, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, x+ w, y + h);
		const shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);
		
		shape.setParent(drawing);
		drawing.Set_GraphicObject(shape);
		drawing.Set_DrawingType(drawing_Anchor);
		drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		drawing.Set_Distance(0, 0, 0, 0);
		logicDocument.AddToParagraph(drawing);
		return drawing;
	}
	function createTextBoxContent(drawing)
	{
		let shape = drawing.GraphicObj;
		shape.createTextBoxContent();
		return shape.getDocContent();
		
	}
	
	
	QUnit.module("Check text input in the document editor");

	let GetParagraphText = AscTest.GetParagraphText;

	QUnit.test("EnterText/CorrectEnterText/CompositeInput", function (assert)
	{
		AscTest.ClearDocument();
		
		let p = new AscWord.Paragraph();
		logicDocument.AddToContent(0, p);
		
		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(), "\r\n", "Check empty selection");
		
		logicDocument.AddTextWithPr("Hello World!");
		
		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false), "Hello World!\r\n", "Add text 'Hello World!'");
		
		logicDocument.MoveCursorToStartPos();
		logicDocument.MoveCursorRight();
		logicDocument.MoveCursorRight();
		
		AscTest.EnterText("123");
		assert.strictEqual(GetParagraphText(p), "He123llo World!", "Add text '123'");
		
		AscTest.EnterText("AA");
		assert.strictEqual(GetParagraphText(p), "He123AAllo World!", "Add text 'AA'");
		
		AscTest.CorrectEnterText("AB", "ABC");
		assert.strictEqual(GetParagraphText(p), "He123AAllo World!", "Check wrong correction AB to ABC");
		
		AscTest.CorrectEnterText("AA", "ABC");
		assert.strictEqual(GetParagraphText(p), "He123ABCllo World!", "Check correction AA to ABC");
		
		AscTest.EnterText("DD");
		logicDocument.MoveCursorLeft();
		AscTest.CorrectEnterText("DD", "CC");
		assert.strictEqual(GetParagraphText(p), "He123ABCDDllo World!", "Add text DD move left and check wrong correction");
		
		logicDocument.MoveCursorToEndPos();
		AscTest.EnterText("qq");
		AscTest.CorrectEnterText("!qq", "!?");
		assert.strictEqual(GetParagraphText(p), "He123ABCDDllo World!?", "Move to the end, add qq and correct !qq to !?");
		
		AscTest.BeginCompositeInput();
		AscTest.ReplaceCompositeInput("WWW");
		AscTest.ReplaceCompositeInput("123");
		AscTest.EndCompositeInput();
		assert.strictEqual(GetParagraphText(p), "He123ABCDDllo World!?123", "Add text '123' with composite input");
		
		AscTest.EnterTextCompositeInput("Zzz");
		AscTest.CorrectEnterText("3Zzz", "$");
		assert.strictEqual(GetParagraphText(p), "He123ABCDDllo World!?12$", "Add text 'Zzz' with composite input and correct it from '3Zzz' to '$'");
		
		AscTest.ClearParagraph(p);
		AscTest.EnterTextCompositeInput("x");
		AscTest.EnterTextCompositeInput("yz");
		AscTest.CorrectEnterText("yz", "x");
		assert.strictEqual(GetParagraphText(p), "xx", "Test special case, when added symbols collect to a single grapheme with previous symbols");
		
		// Вводим на сингальском 1ෑඒ (1-e-e-t)
		p = new AscWord.Paragraph();
		logicDocument.PushToContent(p);
		AscTest.MoveCursorToParagraph(p);
		AscTest.EnterText("1");
		AscTest.BeginCompositeInput();
		AscTest.ReplaceCompositeInput([]);
		AscTest.ReplaceCompositeInput([3536]);
		AscTest.ReplaceCompositeInput([3537]);
		AscTest.ReplaceCompositeInput([3537, 3474]);
		AscTest.EndCompositeInput();
		assert.strictEqual(GetParagraphText(p), "1ෑඒ", "Add text '1ෑඒ'");
		assert.strictEqual(p.IsCursorAtEnd(), true, "Check cursor position");
	});
	QUnit.test("EnterText/CorrectEnterText/CompositeInput in collaboration", function (assert)
	{
		AscTest.StartCollaboration();
		
		AscTest.ClearDocument();
		let p = new AscWord.Paragraph();
		logicDocument.AddToContent(0, p);
		
		AscTest.EnterText("ABC");
		assert.strictEqual(GetParagraphText(p), "ABC", "Add text 'ABC' in collaboration");
		
		AscTest.MoveCursorLeft();
		AscTest.EnterText("111");
		AscTest.SyncCollaboration();
		AscTest.CorrectEnterText("11", "23");
		AscTest.SyncCollaboration();
		assert.strictEqual(GetParagraphText(p), "AB123C", "Add text '111' and correct it with '123' in collaboration (sync between actions)");
		
		AscTest.EnterText("QQQ");
		AscTest.CorrectEnterText("QQ", "RS");
		AscTest.SyncCollaboration();
		assert.strictEqual(GetParagraphText(p), "AB123QRSC", "Add text '111' and correct it with '123' in collaboration (no sync between actions)");
		
		AscTest.EndCollaboration();
	});
	QUnit.test("EnterText/CorrectEnterText/CompositeInput with running TextSpeaker", function (assert)
	{
		AscTest.StartTextSpeaker();
		
		AscTest.ClearDocument();
		let p = new AscWord.Paragraph();
		logicDocument.AddToContent(0, p);
		
		AscTest.EnterText("ABC");
		assert.strictEqual(GetParagraphText(p), "ABC", "Add text 'ABC' in collaboration");
		
		AscTest.MoveCursorLeft();
		AscTest.EnterText("111");
		AscTest.CorrectEnterText("11", "23");
		assert.strictEqual(GetParagraphText(p), "AB123C", "Add text '111' and correct it with '123' in collaboration (sync between actions)");
		
		AscTest.EnterText("QQQ");
		AscTest.CorrectEnterText("QQ", "RS");
		assert.strictEqual(GetParagraphText(p), "AB123QRSC", "Add text '111' and correct it with '123' in collaboration (no sync between actions)");
		
		// Вводим на сингальском 1ෑඒ (1-e-e-t)
		p = new AscWord.Paragraph();
		logicDocument.PushToContent(p);
		AscTest.MoveCursorToParagraph(p);
		AscTest.EnterText("1");
		AscTest.BeginCompositeInput();
		AscTest.ReplaceCompositeInput([]);
		AscTest.ReplaceCompositeInput([3536]);
		AscTest.ReplaceCompositeInput([3537]);
		AscTest.ReplaceCompositeInput([3537, 3474]);
		AscTest.EndCompositeInput();
		assert.strictEqual(GetParagraphText(p), "1ෑඒ", "Add text '1ෑඒ'");
		assert.strictEqual(p.IsCursorAtEnd(), true, "Check cursor position");
		
		AscTest.StopTextSpeaker();
	});
	QUnit.test("Test 'complex script' property on input", function (assert)
	{
		function CheckParagraphSplit(p, text, runCS, runText)
		{
			let count = runCS.length;
			assert.strictEqual(p.GetElementsCount(), count, "Check runs count on entering '" + text + "'");

			for (let index = 0; index < count; ++index)
			{
				let element = p.GetElement(index);

				if (!(element instanceof AscWord.CRun))
				{
					assert.ok(false, "Not a run");
					break;
				}

				assert.strictEqual(element.GetText(), runText[index], `Check run[${index}] text`);
				assert.strictEqual(element.IsCS(), runCS[index], `Check run[${index}] CS`);
			}
		}
		function CheckTextEnter(arrText, arrComposite, ...args)
		{
			AscTest.ClearDocument();

			let p = new AscWord.Paragraph();
			logicDocument.AddToContent(0, p);

			let overallText = "";
			for (let index = 0, count = arrText.length; index < count; ++index)
			{
				let text = arrText[index];
				if (arrComposite[index])
					AscTest.EnterTextCompositeInput(text);
				else
					AscTest.EnterText(text);

				overallText += text;
			}

			CheckParagraphSplit(p, overallText, ...args);
		}


		CheckTextEnter(
			["Abcෑඒ"],
			[false],
			[false, true],
			["Abc", "ෑඒ"]

		);
		CheckTextEnter(
			["Abc1ෑඒ"],
			[false],
			[false, true],
			["Abc1", "ෑඒ"]
		);
		CheckTextEnter(
			["Abc1ෑඒabc"],
			[false],
			[false, true, false],
			["Abc1", "ෑඒ", "abc"]
		);

		// Композитный ввод всегда добавляет новый ран
		CheckTextEnter(
			["Abc", "ෑඒ"],
			[false, true],
			[false, true],
			["Abc", "ෑඒ"]
		);

		CheckTextEnter(
			["Abc", "ෑඒ", "efg"],
			[false, true, false],
			[false, true, false],
			["Abc", "ෑඒ", "efg"]
		);

		// Проверяем, что если ран состоит только из Script_Common, то мы параметр CS подхватывается из следующего ввода
		CheckTextEnter(
			["1abc"],
			[false],
			[false],
			["1abc"]
		);

		CheckTextEnter(
			["2ෑඒ"],
			[false],
			[true],
			["2ෑඒ"]
		);

		CheckTextEnter(
			["3", "ෑඒ"],
			[false, true],
			[true, true],
			["3", "ෑඒ"]
		);
	});
	QUnit.test("Test EnterText/CorrectEnterText/CompositeInput in Shape", function (assert)
	{
		// Делаем тест по сценарию бага 67336 для прокерки корректности ввода в автофигуру
		AscTest.ClearDocument();
		
		let p = new AscWord.Paragraph();
		logicDocument.AddToContent(0, p);
		
		let drawing = addShape(0, 0, 100, 100);
		p.AddToParagraph(drawing);
		
		let docContent = createTextBoxContent(drawing);
		pInShape = docContent.GetElement(0);
		
		assert.strictEqual(GetParagraphText(p), "");
		assert.strictEqual(GetParagraphText(pInShape), "");
		
		AscTest.MoveCursorToParagraph(pInShape, true);

		let _textInShape = "Text in shape";
		AscTest.BeginCompositeInput();
		AscTest.ReplaceCompositeInput("Test");
		AscTest.ReplaceCompositeInput(_textInShape);
		AscTest.EndCompositeInput();
		assert.strictEqual(GetParagraphText(p), "");
		assert.strictEqual(GetParagraphText(pInShape), _textInShape);

		AscTest.CorrectEnterText(_textInShape, "");
		assert.strictEqual(GetParagraphText(p), "");
		assert.strictEqual(GetParagraphText(pInShape), "");

		AscTest.EnterText(_textInShape);
		assert.strictEqual(GetParagraphText(p), "");
		assert.strictEqual(GetParagraphText(pInShape), _textInShape);
	});
});
