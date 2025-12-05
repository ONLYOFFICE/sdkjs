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
		
		logicDocument.RemoveFromContent(1, 1);
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				{"FieldId" : "3", "Value" : "Addin №2", "Content" : "This is the second addin field"}
			],
			"Remove the paragraph with the first field and check get addin function");
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
	
	
});
