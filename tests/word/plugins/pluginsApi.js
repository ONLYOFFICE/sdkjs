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
	
	let PluginsApi = AscTest.Editor;
	
	let logicDocument = AscTest.CreateLogicDocument();
	logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);
	
	function MoveToNewParagraph()
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(logicDocument.GetElementsCount(), p);
		p.SetThisElementCurrent();
		return p;
	}
	
	QUnit.module("Test plugins api");
	
	QUnit.test("Test work with addin fields", function (assert)
	{
		MoveToNewParagraph();
		
		assert.strictEqual(PluginsApi.pluginMethod_GetAllAddinFields().length, 0, "Check addin fields in empty document");
		
		MoveToNewParagraph();
		PluginsApi.pluginMethod_AddAddinField({"Value" : "Test addin", "Content" : 123});
		
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				{"FieldId" : "1", "Value" : "Test addin", "Content" : "123"}
			],
			"Add addin field and check get function");
		
		MoveToNewParagraph();
		assert.strictEqual(logicDocument.GetAllFields().length, 1, "Check the number of all fields");
		logicDocument.AddFieldWithInstruction("PAGE");
		assert.strictEqual(logicDocument.GetAllFields().length, 2, "Add PAGE field and check the number of all fields");
		
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields().length, 1, "Check the number of addin fields");
		
		MoveToNewParagraph();
		PluginsApi.pluginMethod_AddAddinField({"Value" : "Addin №2", "Content" : "This is the second addin field"});
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				{"FieldId" : "1", "Value" : "Test addin", "Content" : "123"},
				{"FieldId" : "3", "Value" : "Addin №2", "Content" : "This is the second addin field"}
			],
			"Add addin field and check get function");
		
		PluginsApi.pluginMethod_UpdateAddinFields(
			[
				{"FieldId" : "1", "Value" : "Addin №1", "Content" : "This is the first addin field"},
			],
			"Add addin field and check get function");
		
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				{"FieldId" : "1", "Value" : "Addin №1", "Content" : "This is the first addin field"},
				{"FieldId" : "3", "Value" : "Addin №2", "Content" : "This is the second addin field"}
			],
			"Change the first adding and check get function");
		
		PluginsApi.pluginMethod_SelectAddinField("1");
		assert.strictEqual(logicDocument.GetSelectedText(), "This is the first addin field", "Check addin field selection");
		
		PluginsApi.pluginMethod_SelectAddinField("25");
		assert.strictEqual(logicDocument.GetSelectedText(), "This is the first addin field", "Check addin field selection");
		
		PluginsApi.pluginMethod_SelectAddinField("3");
		assert.strictEqual(logicDocument.GetSelectedText(), "This is the second addin field", "Check addin field selection");
		
		logicDocument.RemoveFromContent(1, 1);
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				{"FieldId" : "3", "Value" : "Addin №2", "Content" : "This is the second addin field"}
			],
			"Remove the paragraph with the first field and check get addin function");
		
		PluginsApi.pluginMethod_RemoveAddinField("3");
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[],
			"Remove the first add-in field");
		
	});

	QUnit.test("Test addin fields in header/footer", function (assert)
	{
		AscTest.ClearDocument();
		MoveToNewParagraph();

		let sectPr = logicDocument.GetFinalSectPr();
		let header = AscTest.CreateDefaultHeader(sectPr);
		header.Set_CurrentElement(false, 0);
		
		logicDocument.Set_DocumentDefaultTab(35.45);

		let addinFieldData = {"FieldId": "1", "Value": "Addin №1", "Content": "This is the first addin field"};
		PluginsApi.pluginMethod_AddAddinField(addinFieldData);

		logicDocument.RemoveSelection();
		logicDocument.Document_UpdateInterfaceState();

		let fields = PluginsApi.pluginMethod_GetAllAddinFields();
		assert.strictEqual(fields.length, 1, "Check one addin field in header");
		assert.strictEqual(fields[0].Value, "Addin №1", "Check Value");
		assert.strictEqual(fields[0].Content, "This is the first addin field", "Check Content from header field");
	});

	QUnit.test("Test RemoveFieldWrapper", function(assert)
	{
		AscTest.ClearDocument();
		MoveToNewParagraph();
		assert.strictEqual(logicDocument.GetAllFields().length, 0, "Check the number of all fields in the empty document");
		
		MoveToNewParagraph();
		logicDocument.AddFieldWithInstruction("PAGE");
		
		let p = MoveToNewParagraph();
		let field = logicDocument.AddFieldWithInstruction("PAGE");
		assert.strictEqual(logicDocument.GetAllFields().length, 2, "Add two PAGE fields and check count of all fields");
		
		logicDocument.UpdateFields(false);
		
		assert.strictEqual(AscTest.GetParagraphText(p), "1", "Check the text of the third paragraph");
		
		let fieldId = field.GetFieldId();
		PluginsApi.pluginMethod_RemoveFieldWrapper(fieldId);
		assert.strictEqual(logicDocument.GetAllFields().length, 1, "Remove field wrapper from second field and check number of fields");
		assert.strictEqual(AscTest.GetParagraphText(p), "1", "Check the text of the third paragraph");
	});
	
	QUnit.test("Test SetEditingRestrictions", function(assert)
	{
		AscTest.ClearDocument();
		MoveToNewParagraph();
		
		assert.strictEqual(logicDocument.CanEdit(), true, "Check if we can edit new document");
		
		PluginsApi.pluginMethod_SetEditingRestrictions("readOnly");
		assert.strictEqual(logicDocument.CanEdit(), false, "Set read only restriction and check if we can edit document");
		
		// Set to none to pass subsequent tests
		PluginsApi.pluginMethod_SetEditingRestrictions("none");
	});
	
	QUnit.test("Test CurrenWord/CurrentSentence", function(assert)
	{
		AscTest.ClearDocument();
		let p = MoveToNewParagraph();
		AscTest.EnterText("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.    Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.");
		
		AscTest.MoveCursorToParagraph(p, true);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentWord(), "Lorem", "Check current word at the start of the paragraph");
		AscTest.MoveCursorRight(false, false, 6);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentWord(), "ipsum", "Move cursor right(6) and check current word on the left edge of the word");
		AscTest.MoveCursorRight(false, false, 5);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentWord(), "ipsum", "Move cursor right(5) and check current word on the right edge of the word");
		AscTest.MoveCursorToParagraph(p, false);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentWord(), ".", "Check current word at the end of the paragraph");
		AscTest.MoveCursorLeft(false, false, 1);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentWord(), "laborum", "Move cursor left and check current word");
		
		AscTest.MoveCursorToParagraph(p, true);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(),
			"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
			"Check current sentence at the start of the paragraph");
		
		AscTest.MoveCursorToParagraph(p, false);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(),
			"",
			"Check current sentence at the end of the paragraph");
		
		AscTest.MoveCursorLeft(false, false, 5);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(),
			"Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.",
			"Move cursor left(5) and check current sentence");
		
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 123);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(),
			"Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.",
			"Move to the start of the second sentence and check it");
		
		AscTest.ClearDocument();
		p = MoveToNewParagraph();
		AscTest.EnterText("Test text");
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 2);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentWord(), "Test", "Add new paragraph and check current word");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(), "Test text", "Check current sentence");
		
		logicDocument.AddFieldWithInstruction("PAGE");
		AscTest.Recalculate();
		AscTest.MoveCursorToParagraph(p, true);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentWord(), "Te", "Add hidden complex field in the middle of word 'Test' and check current word");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(), "Te1st text", "Check current sentence");
		
		
		AscTest.ClearDocument();
		p = MoveToNewParagraph();
		AscTest.EnterText("Test text");

		AscTest.MoveCursorToParagraph(p, true);
		PluginsApi.pluginMethod_ReplaceCurrentWord("First");
		assert.strictEqual(AscTest.GetParagraphText(p), "First text", "Replace current word at the start of the paragraph");
		
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 2);
		PluginsApi.pluginMethod_ReplaceCurrentWord("Second");
		assert.strictEqual(AscTest.GetParagraphText(p), "Second text", "Replace current word at the second position of the paragraph");
		
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 3);
		PluginsApi.pluginMethod_ReplaceCurrentWord("123", "afterCursor");
		assert.strictEqual(AscTest.GetParagraphText(p), "Sec123 text", "Replace the part of the word after cursor");
		
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 3);
		PluginsApi.pluginMethod_ReplaceCurrentWord("654", "beforeCursor");
		assert.strictEqual(AscTest.GetParagraphText(p), "654123 text", "Replace the part of the word before cursor");
		
		
		AscTest.ClearDocument();
		p = MoveToNewParagraph();
		AscTest.EnterText("The quick brown fox jumps over the lazy dog. The five boxing wizards jump quickly. Eat more of those fresh french loafs and drink a tea!");
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 16);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("entirely"), "The quick brown fox jumps over the lazy dog.", "Check current sentence");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("afterCursor"), "fox jumps over the lazy dog.", "Check the right part of the current sentence");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("beforeCursor"), "The quick brown ", "Check the left part of the current sentence");
		
		AscTest.MoveCursorRight(false, false, 28);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("entirely"), "The five boxing wizards jump quickly.", "Check current sentence");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("afterCursor"), "The five boxing wizards jump quickly.", "Check the right part of the current sentence");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("beforeCursor"), "", "Check the left part of the current sentence");
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(false, false, 1);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("entirely"), "Eat more of those fresh french loafs and drink a tea!", "Check current sentence");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("afterCursor"), "!", "Check the right part of the current sentence");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence("beforeCursor"), "Eat more of those fresh french loafs and drink a tea", "Check the left part of the current sentence");
		
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 16);
		PluginsApi.pluginMethod_ReplaceCurrentSentence("The slow yellow rabbit jumps over the fluffy cat!", "entirely");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(), "The five boxing wizards jump quickly.", "Replace first sentence and check next sentence.");
		AscTest.MoveCursorLeft(false, false, 5);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(), "The slow yellow rabbit jumps over the fluffy cat!", "Check replaced sentence.");
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 58);
		PluginsApi.pluginMethod_ReplaceCurrentSentence("The eight", "beforeCursor");
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(), "The eight boxing wizards jump quickly.", "Replace left part of the sentence.");
		PluginsApi.pluginMethod_ReplaceCurrentSentence(" relaxing wizards jump slowly.", "afterCursor");
		AscTest.MoveCursorLeft(false, false, 5);
		assert.strictEqual(PluginsApi.pluginMethod_GetCurrentSentence(), "The eight relaxing wizards jump slowly.", "Replace right part of the sentence.");
		
		AscTest.ClearDocument();
		p = MoveToNewParagraph();
		AscTest.EnterText("The quick brown fox jumps over the lazy dog. The five boxing wizards jump quickly. Eat more of those fresh french loafs and drink a tea!");
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 64);
		PluginsApi.pluginMethod_ReplaceCurrentSentence("The five boxing wizards jump quickly.", "entirely");
		assert.strictEqual(AscTest.GetParagraphText(p), "The quick brown fox jumps over the lazy dog. The five boxing wizards jump quickly. Eat more of those fresh french loafs and drink a tea!", "Replace sentence on itself and check all text");
		
		
	})

	QUnit.test("Test ApiRun.GetInlineDrawings", function (assert)
	{
		// --- Test 1: Run with text only, no drawings ---
		AscTest.ClearDocument();
		let p = MoveToNewParagraph();
		AscTest.EnterText("Hello World");

		let apiPara = Api.CreateParagraph();
		logicDocument.AddToContent(logicDocument.GetElementsCount(), apiPara.Paragraph);
		apiPara.Paragraph.SetThisElementCurrent();

		// Get the run that contains "Hello World" from the first paragraph
		let p1 = logicDocument.GetElement(logicDocument.GetElementsCount() - 2);
		let apiPara1 = new AscWord.ApiParagraph(p1);
		let run1 = apiPara1.GetElement(0);

		assert.strictEqual(run1.GetClassType(), "run", "Element is a run");
		let result1 = run1.GetInlineDrawings();
		assert.ok(Array.isArray(result1), "GetInlineDrawings returns an array");
		assert.strictEqual(result1.length, 0, "Text-only run returns empty array");

		// --- Test 2: Run with a drawing only (no text) ---
		AscTest.ClearDocument();
		p = MoveToNewParagraph();

		let internalRun2 = new ParaRun(p, false);
		p.AddToContentToEnd(internalRun2);

		let drawing2 = AscTest.CreateImage(50, 50);
		internalRun2.Add_ToContent(0, drawing2);
		drawing2.Set_Parent(internalRun2);

		let apiPara2 = new AscWord.ApiParagraph(p);
		let apiRun2 = apiPara2.GetElement(0);
		let result2 = apiRun2.GetInlineDrawings();

		assert.strictEqual(result2.length, 1, "Drawing-only run has 1 inline drawing");
		assert.strictEqual(result2[0]["position"], 0, "Drawing-only run: position is 0");
		assert.ok(result2[0]["drawing"], "Drawing-only run: drawing object exists");

		// --- Test 3: Drawing between text ("AB[img]CD") ---
		AscTest.ClearDocument();
		p = MoveToNewParagraph();

		let internalRun3 = new ParaRun(p, false);
		internalRun3.AddText("ABCD");
		p.AddToContentToEnd(internalRun3);

		let drawing3 = AscTest.CreateImage(50, 50);
		internalRun3.Add_ToContent(2, drawing3); // After A, B
		drawing3.Set_Parent(internalRun3);

		let apiPara3 = new AscWord.ApiParagraph(p);
		let apiRun3 = apiPara3.GetElement(0);
		let result3 = apiRun3.GetInlineDrawings();
		let text3 = apiRun3.GetText();

		assert.strictEqual(result3.length, 1, "Drawing between text: 1 inline drawing");
		assert.strictEqual(result3[0]["position"], 2, "Drawing between text: position is 2 (after AB)");
		assert.strictEqual(text3, "ABCD", "GetText returns text without drawing");

		// Verify text splitting works correctly
		let before3 = text3.substring(0, result3[0]["position"]);
		let after3 = text3.substring(result3[0]["position"]);
		assert.strictEqual(before3, "AB", "Text before drawing is 'AB'");
		assert.strictEqual(after3, "CD", "Text after drawing is 'CD'");

		// --- Test 4: Drawing at start ("[img]ABCD") ---
		AscTest.ClearDocument();
		p = MoveToNewParagraph();

		let internalRun4 = new ParaRun(p, false);
		internalRun4.AddText("ABCD");
		p.AddToContentToEnd(internalRun4);

		let drawing4 = AscTest.CreateImage(50, 50);
		internalRun4.Add_ToContent(0, drawing4); // At start
		drawing4.Set_Parent(internalRun4);

		let apiPara4 = new AscWord.ApiParagraph(p);
		let apiRun4 = apiPara4.GetElement(0);
		let result4 = apiRun4.GetInlineDrawings();

		assert.strictEqual(result4.length, 1, "Drawing at start: 1 inline drawing");
		assert.strictEqual(result4[0]["position"], 0, "Drawing at start: position is 0");

		// --- Test 5: Drawing at end ("ABCD[img]") ---
		AscTest.ClearDocument();
		p = MoveToNewParagraph();

		let internalRun5 = new ParaRun(p, false);
		internalRun5.AddText("ABCD");
		p.AddToContentToEnd(internalRun5);

		let drawing5 = AscTest.CreateImage(50, 50);
		internalRun5.Add_ToContent(internalRun5.Content.length, drawing5); // At end
		drawing5.Set_Parent(internalRun5);

		let apiPara5 = new AscWord.ApiParagraph(p);
		let apiRun5 = apiPara5.GetElement(0);
		let result5 = apiRun5.GetInlineDrawings();

		assert.strictEqual(result5.length, 1, "Drawing at end: 1 inline drawing");
		assert.strictEqual(result5[0]["position"], 4, "Drawing at end: position equals text length (4)");

		// --- Test 6: Multiple drawings ("A[img1]B[img2]C") ---
		AscTest.ClearDocument();
		p = MoveToNewParagraph();

		let internalRun6 = new ParaRun(p, false);
		internalRun6.AddText("ABC");
		p.AddToContentToEnd(internalRun6);

		// Insert second drawing first (at higher index) to avoid shifting
		let drawing6b = AscTest.CreateImage(50, 50);
		internalRun6.Add_ToContent(2, drawing6b); // After B: "AB[img2]C"
		drawing6b.Set_Parent(internalRun6);

		let drawing6a = AscTest.CreateImage(50, 50);
		internalRun6.Add_ToContent(1, drawing6a); // After A: "A[img1]B[img2]C"
		drawing6a.Set_Parent(internalRun6);

		let apiPara6 = new AscWord.ApiParagraph(p);
		let apiRun6 = apiPara6.GetElement(0);
		let result6 = apiRun6.GetInlineDrawings();

		assert.strictEqual(result6.length, 2, "Multiple drawings: 2 inline drawings");
		assert.strictEqual(result6[0]["position"], 1, "Multiple drawings: first at position 1");
		assert.strictEqual(result6[1]["position"], 2, "Multiple drawings: second at position 2");

		// --- Test 7: Verify position matches GetText() index ---
		AscTest.ClearDocument();
		p = MoveToNewParagraph();

		let internalRun7 = new ParaRun(p, false);
		internalRun7.AddText("Hi ");
		p.AddToContentToEnd(internalRun7);

		let drawing7 = AscTest.CreateImage(50, 50);
		internalRun7.Add_ToContent(internalRun7.Content.length, drawing7);
		drawing7.Set_Parent(internalRun7);

		// Add more text after drawing in the same run
		let textAfter = new AscWord.CRunText(0x0057); // 'W'
		internalRun7.Add_ToContent(internalRun7.Content.length, textAfter);
		let textAfter2 = new AscWord.CRunText(0x006F); // 'o'
		internalRun7.Add_ToContent(internalRun7.Content.length, textAfter2);

		let apiPara7 = new AscWord.ApiParagraph(p);
		let apiRun7 = apiPara7.GetElement(0);
		let text7 = apiRun7.GetText();
		let result7 = apiRun7.GetInlineDrawings();

		assert.strictEqual(text7, "Hi Wo", "GetText returns 'Hi Wo' (3 + 2 chars)");
		assert.strictEqual(result7.length, 1, "One drawing found");
		assert.strictEqual(result7[0]["position"], 3, "Drawing at position 3 (after 'Hi ')");

		// Verify splitting produces correct segments
		let before7 = text7.substring(0, result7[0]["position"]);
		let after7 = text7.substring(result7[0]["position"]);
		assert.strictEqual(before7, "Hi ", "Text before drawing is 'Hi '");
		assert.strictEqual(after7, "Wo", "Text after drawing is 'Wo'");
	});


});
