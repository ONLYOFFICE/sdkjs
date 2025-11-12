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
	let styleManager  = logicDocument.GetStyleManager();
	
	let paraStyle1 = AscTest.CreateParagraphStyle("ParaStyle1");
	let paraStyle2 = AscTest.CreateParagraphStyle("ParaStyle2");
	
	let runStyle1 = AscTest.CreateRunStyle("RunStyle1");
	let runStyle2 = AscTest.CreateRunStyle("RunStyle2");
	
	let defaultPStyle = styleManager.Get(styleManager.GetDefaultParagraph());
	let noStyle       = "";
	
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
	function SetPStyle(p, pStyle)
	{
		p.SetPStyle(pStyle ? pStyle.GetId() : null);
	}
	function SetRStyle(run, rStyle)
	{
		run.SetRStyle(rStyle ? rStyle.GetId() : null);
	}
	
	let displayStyleName = "";
	editor.isDocumentLoadComplete = true;
	editor.sync_ParaStyleName = function(styleName)
	{
		displayStyleName = styleName;
	};
	
	function CheckStyle(assert, expectedStyle, message)
	{
		logicDocument.UpdateInterface();
		assert.strictEqual(displayStyleName, expectedStyle ? expectedStyle.GetName() : "", message ? message : "");
	}

	QUnit.module("Cursor");
	
	QUnit.test("Run with/without style in paragraph", function(assert)
	{
		let word = "Word!";
		
		AscTest.ClearDocument();
		let p = AddParagraph(0);
		let run = CreateRun(word);
		p.AddToContentToEnd(run);
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(false, false, 1);
		
		SetPStyle(p, paraStyle1);
		SetRStyle(run, runStyle1);
		
		CheckStyle(assert, runStyle1, "Move cursor to the last letter of the run with a style");
		run.SetRStyle(null);
		CheckStyle(assert, paraStyle1, "Remove style from run");
		p.SetPStyle(null);
		CheckStyle(assert, defaultPStyle, "Remove style from paragraph");
	});

	QUnit.module("Select");
	
	QUnit.test("Two different paragraphs", function(assert)
	{
		let word1 = "Word!";
		let word2 = "Hello!";
		
		AscTest.ClearDocument();
		
		let p1 = AddParagraph(0);
		let run1 = CreateRun(word1);
		p1.AddToContentToEnd(run1);
		
		let p2 = AddParagraph(1);
		let run2 = CreateRun(word2);
		p2.AddToContentToEnd(run2);
		AscTest.SelectDocumentRange(0, 1);
		
		CheckStyle(assert, defaultPStyle, "Both paragraph have no style (no run style)");
		
		SetPStyle(p1, paraStyle1);
		SetPStyle(p2, paraStyle1);
		CheckStyle(assert, paraStyle1, "Both paragraphs have the same style ParaStyle1 (no run style)");
		
		SetPStyle(p1, paraStyle1);
		SetPStyle(p2, paraStyle2);
		CheckStyle(assert, noStyle, "Paragraphs have different styles (no run style)");
		
		SetRStyle(run1, runStyle1);
		SetRStyle(run2, runStyle2);
		
		logicDocument.RemoveSelection();
		AscTest.MoveCursorToParagraph(p2, false);
		AscTest.MoveCursorLeft(true, false, word2.length - 1);
		CheckStyle(assert, runStyle2, "Apply RunStyle1 to run1 and RunStyle2 to run2. Partially select the second run");
		
		SetRStyle(run1, null);
		SetRStyle(run2, null);
		CheckStyle(assert, paraStyle2, "Remove style from runs. Leave partially select");
		
		AscTest.SelectDocumentRange(1, 1);
		CheckStyle(assert, paraStyle2, "Select entire second paragraph");
		
		logicDocument.RemoveSelection();
		AscTest.MoveCursorToParagraph(p2, false);
		AscTest.MoveCursorLeft(true, false, word1.length + word2.length);
		CheckStyle(assert, noStyle, "Select both paragraph by key arrows");
	});
	
	QUnit.test("Different runs in one paragraph", function(assert)
	{
		let word1 = "Word!";
		let word2 = "Hello!";
		
		AscTest.ClearDocument();
		
		let p = AddParagraph(0);
		let run1 = CreateRun(word1);
		let run2 = CreateRun(word2);
		
		p.AddToContentToEnd(run1);
		p.AddToContentToEnd(run2);
		
		SetPStyle(p, paraStyle1);
		SetRStyle(run1, runStyle1);
		SetRStyle(run2, runStyle1);
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(true, false, word2.length);
		CheckStyle(assert, runStyle1, "Add two runs with different styles in one paragraph and select second run");
		
		AscTest.MoveCursorLeft(true, false, word1.length);
		CheckStyle(assert, runStyle1, "Add first run to selection");
		
		SetRStyle(run2, runStyle2);
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(true, false, word2.length);
		CheckStyle(assert, runStyle2, "Set the style of run2 to RunStyle2 and select second run");
		
		AscTest.MoveCursorLeft(true, false, word1.length);
		CheckStyle(assert, paraStyle1, "Add first run to selection (runs with different styles in the selection)");
	});
	
	QUnit.test("Multiple runs with same style in different paragraphs", function(assert)
	{
		let word1 = "Word!";
		let word2 = "NextWord";
		let word3 = "Hello!";
		
		AscTest.ClearDocument();
		
		let p1 = AddParagraph(0);
		let run1 = CreateRun(word1);
		let run2 = CreateRun(word2);
		p1.AddToContentToEnd(run1);
		p1.AddToContentToEnd(run2);
		
		let p2 = AddParagraph(1);
		let run3 = CreateRun(word3);
		p2.AddToContentToEnd(run3);
		
		SetPStyle(p1, paraStyle1);
		SetPStyle(p2, paraStyle2);
		
		SetRStyle(run1, runStyle1);
		SetRStyle(run2, runStyle2);
		SetRStyle(run3, runStyle2);
		
		AscTest.MoveCursorToParagraph(p2, false);
		AscTest.MoveCursorLeft(true, false, word3.length + word2.length);
		CheckStyle(assert, runStyle2, "Select two last runs with same style in different paragraphs with different styles");
		
		AscTest.MoveCursorLeft(true, false, word1.length);
		CheckStyle(assert, noStyle, "Add the first run to selection");
		
		SetPStyle(p1, paraStyle1);
		SetPStyle(p2, paraStyle1);
		CheckStyle(assert, paraStyle1, "Make both paragraphs the same style");
		
		SetPStyle(p1, null);
		SetPStyle(p2, null);
		CheckStyle(assert, defaultPStyle, "Remove style from both paragraphs");
	});
	
	QUnit.test("Combination of spaces and text", function(assert)
	{
		// Если есть хоть 1 текстовый элемент в выделении, то используем его стиль.
		// В противном случае используем первый попавшийся стиль
		// Если стиля нет, то используем уже стиль параграфа
		
		let word1 = "Word!";
		let space1 = "   ";
		let space2 = "   ";
		let space3 = "   ";
		let word2 = "Hello!";
		
		AscTest.ClearDocument();
		
		let p = AddParagraph(0);
		let textRun1 = CreateRun(word1);
		let textRun2 = CreateRun(word2);
		let spaceRun1 = CreateRun(space1);
		let spaceRun2 = CreateRun(space2);
		let spaceRun3 = CreateRun(space3);
		
		p.AddToContentToEnd(textRun1);
		p.AddToContentToEnd(spaceRun1);
		p.AddToContentToEnd(spaceRun2);
		p.AddToContentToEnd(spaceRun3);
		p.AddToContentToEnd(textRun2);
		
		SetPStyle(p, paraStyle1);
		SetRStyle(textRun1, runStyle1);
		SetRStyle(spaceRun1, null);
		SetRStyle(spaceRun2, runStyle2);
		SetRStyle(spaceRun3, runStyle1);
		SetRStyle(textRun2, runStyle1);
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(true, false, word2.length);
		CheckStyle(assert, runStyle1, "Select second word");
		
		AscTest.MoveCursorLeft(true, false, space3.length);
		CheckStyle(assert, runStyle1, "Add last run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, space2.length);
		CheckStyle(assert, runStyle1, "Add second run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, space1.length);
		CheckStyle(assert, runStyle1, "Add first run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, word1.length);
		CheckStyle(assert, runStyle1, "Add the first run with text to selection");
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(false, false, word2.length);
		AscTest.MoveCursorLeft(true, false, space3.length);
		CheckStyle(assert, runStyle1, "Add last run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, space2.length);
		CheckStyle(assert, runStyle2, "Add second run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, space1.length);
		CheckStyle(assert, paraStyle1, "Add first run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, word1.length);
		CheckStyle(assert, runStyle1, "Add the first run with text to selection");
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(false, false, word2.length + space3.length);
		AscTest.MoveCursorLeft(true, false, space2.length);
		CheckStyle(assert, runStyle2, "Add second run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, space1.length);
		CheckStyle(assert, paraStyle1, "Add first run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, word1.length);
		CheckStyle(assert, runStyle1, "Add the first run with text to selection");
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(false, false, word2.length + space3.length + space2.length);
		AscTest.MoveCursorLeft(true, false, space1.length);
		CheckStyle(assert, paraStyle1, "Add first run with spaces to selection");
		
		AscTest.MoveCursorLeft(true, false, word1.length);
		CheckStyle(assert, runStyle1, "Add the first run with text to selection");
		
		// Check simple cursor position
		
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 1);
		CheckStyle(assert, runStyle1, "Move to 1 position in the first word");
		
		AscTest.MoveCursorRight(false, false, word1.length);
		CheckStyle(assert, paraStyle1, "Move to 1 position in the first space run");
		
		AscTest.MoveCursorRight(false, false, space1.length);
		CheckStyle(assert, runStyle2, "Move to 1 position in the second space run");
		
		AscTest.MoveCursorRight(false, false, space2.length);
		CheckStyle(assert, runStyle1, "Move to 1 position in the last space run");
		
		AscTest.MoveCursorRight(false, false, space3.length);
		CheckStyle(assert, runStyle1, "Move to 1 position in the last word");
	});
});
