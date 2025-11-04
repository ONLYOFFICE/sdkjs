/*
 * (c) Copyright Ascensio System SIA 2010-2025
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
	QUnit.module("Test the ApiSelection methods");

	AscTest.Editor.sync_AddComment = AscCommon.DocumentEditorApi.prototype.sync_AddComment.bind(AscTest.Editor);
	AscTest.Editor.sync_RemoveComment = AscCommon.DocumentEditorApi.prototype.sync_RemoveComment.bind(AscTest.Editor);
	AscTest.Editor.private_CreateApiChart = AscCommon.DocumentEditorApi.prototype.private_CreateApiChart.bind(AscTest.Editor);
	AscTest.Editor.CreateNoFill = AscCommon.DocumentEditorApi.prototype.CreateNoFill.bind(AscTest.Editor);
	AscTest.Editor.CreateStroke = AscCommon.DocumentEditorApi.prototype.CreateStroke.bind(AscTest.Editor);

	function createParagraphWithText(text)
	{
		const oParagraph = AscTest.Editor.CreateParagraph();
		oParagraph.AddText(text);
		return oParagraph;
	}
	
	function selectAllContent()
	{
		const oDocument = AscTest.Editor.GetDocument();
		const oRange = oDocument.GetRange();
		if (oRange)
		{
			oRange.Select();
		}
	}
	
	QUnit.test("GetRange", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oPara1 = createParagraphWithText("First paragraph");
		const oPara2 = createParagraphWithText("Second paragraph");
		oDocument.Push(oPara1);
		oDocument.Push(oPara2);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const oRange = oSelection.GetRange();
		
		assert.ok(oRange !== null, "GetRange must return a range object");
		assert.strictEqual(oRange.GetClassType(), "range", "Muust be a range");
		
		const rangeText = oRange.GetText();
		assert.strictEqual(rangeText, "First paragraph\r\nSecond paragraph\r\n", "Range text should match selection");
		
		const paragraphs = oRange.GetAllParagraphs();
		assert.strictEqual(paragraphs.length, 2, "Range should contain 2 paragraphs");
	});
	
	QUnit.test("GetTextPr (Font)", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oParagraph = createParagraphWithText("Test text");
		oDocument.Push(oParagraph);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const oTextPr = oSelection.GetTextPr();
		
		assert.ok(oTextPr !== null, "Check existence");
		assert.strictEqual(oTextPr.GetClassType(), "textPr", "Must be a textPr object");
		
		oTextPr.SetBold(true);
		assert.strictEqual(oTextPr.GetBold(), true, "Bold must be true after SetBold(true)");
		
		oTextPr.SetFontSize(48);
		assert.strictEqual(oTextPr.GetFontSize(), 48, "Font size must be 48 half-points (24pt)");
		
		const oTextPr2 = oSelection.Font;
		assert.strictEqual(oTextPr2.GetBold(), true, "Bold must still be true");
		assert.strictEqual(oTextPr2.GetFontSize(), 48, "Font size must still be 48");
	});
	
	QUnit.test("GetShading", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oParagraph = createParagraphWithText('Test shading');
		oDocument.Push(oParagraph);
		
		const oRun = oParagraph.GetElement(0);
		const oTextPr = oRun.GetTextPr();
		oTextPr.SetShd("clear", 255, 0, 0);
		
		oRun.Select();
		
		const oSelection = oDocument.GetSelection();
		const oShading = oSelection.GetShading();
		
		assert.ok(oShading !== null, "Shading should exist on selected run");
		assert.strictEqual(oShading.GetClassType(), "rgbColor", "Should be an RGB color");
		assert.strictEqual(oShading.GetRGB(), 255 << 16 | 0 << 8 | 0, "Must be red");
	});
	
	QUnit.test("GetStyle", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		
		// Paragraph style
		const oParagraph1 = createParagraphWithText("Test text");
		oDocument.Push(oParagraph1);
		
		selectAllContent();
		let oSelection = oDocument.GetSelection();
		
		let oStyle = oSelection.GetStyle();
		assert.ok(oStyle !== null, "Paragraph style should exist");
		assert.strictEqual(oStyle.GetClassType(), "style", "Should be a style object");
		assert.strictEqual(oStyle.GetName(), "Normal", "Paragraph style name should be 'Normal'");
		assert.strictEqual(oStyle.GetType(), "paragraph", "Should be a paragraph style");
		
		// Character (run) style
		const oParagraph2 = createParagraphWithText("Text with character style");
		oDocument.Push(oParagraph2);
		
		const oCharStyle = oDocument.CreateStyle("TestCharStyle", "run");
		oCharStyle.GetTextPr().SetBold(true);
		
		const oRun = oParagraph2.GetElement(0);
		oRun.SetStyle(oCharStyle);
		oParagraph2.AddElement(oRun);
		
		oRun.Select();
		
		oSelection = oDocument.GetSelection();
		oStyle = oSelection.GetStyle();
		
		assert.ok(oStyle !== null, "Check run style existence");
		assert.strictEqual(oStyle.GetClassType(), "style", "Must be a style object");
		assert.strictEqual(oStyle.GetName(), "TestCharStyle", "Character style name should be 'TestCharStyle'");
		assert.strictEqual(oStyle.GetType(), "run", "Should be a run (character) style");
	});
	
	QUnit.test("GetParaPr (ParagraphFormat)", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oParagraph = createParagraphWithText("Test paragraph");
		oDocument.Push(oParagraph);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const oParaPr = oSelection.GetParaPr();
		oParaPr.SetJc("center");
		
		assert.ok(oParaPr !== null, "GetParaPr should return paragraph properties");
		assert.strictEqual(oParaPr.GetClassType(), "paraPr", "Should be a paraPr object");
		assert.strictEqual(oParaPr.GetJc(), "center", "Alignment should be center");
		
		const oParaPr2 = oSelection.GetParaPr();
		assert.strictEqual(oParaPr2.GetJc(), "center", "Alignment should still be center");
	});
	
	QUnit.test("GetText", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oPara1 = createParagraphWithText("First paragraph");
		const oPara2 = createParagraphWithText("Second paragraph");
		oDocument.Push(oPara1);
		oDocument.Push(oPara2);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const text = oSelection.GetText();
		assert.strictEqual(text, "First paragraph\r\nSecond paragraph\r\n", "Selection text should match the combined paragraph texts");
	});
	
	QUnit.test("GetWords", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oPara1 = createParagraphWithText("First paragraph");
		const oPara2 = createParagraphWithText("Second paragraph");
		oDocument.Push(oPara1);
		oDocument.Push(oPara2);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const words = oSelection.GetWords();
		
		assert.ok(words.length >= 4, "Must return 4 words from two paragraphs");
		assert.strictEqual(words[0].GetText(), "First", "First word must be 'First'");
		assert.strictEqual(words[1].GetText(), "paragraph", "Second word must be 'paragraph'");
		assert.strictEqual(words[2].GetText(), "Second", "Third word must be 'Second'");
		assert.strictEqual(words[3].GetText(), "paragraph", "Fourth word must be 'paragraph'");
	});
	
	QUnit.test("GetCharacters - single paragraph", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oParagraph = createParagraphWithText("AB C");
		oDocument.Push(oParagraph);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const chars = oSelection.GetCharacters();
		
		assert.ok(chars.length >= 4, "Must return at least 4 characters");
		assert.strictEqual(chars[0].GetText(), "A", "First character must be 'A'");
		assert.strictEqual(chars[1].GetText(), "B", "Second character must be 'B'");
		assert.strictEqual(chars[2].GetText(), " ", "Third character must be a space");
		assert.strictEqual(chars[3].GetText(), "C", "Fourth character must be 'C'");
	});
	
	QUnit.test("GetCharacters - multiple paragraphs", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oPara1 = createParagraphWithText("AB");
		const oPara2 = createParagraphWithText("CD");
		oDocument.Push(oPara1);
		oDocument.Push(oPara2);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const chars = oSelection.GetCharacters();
		
		assert.ok(chars.length >= 4, "Must return at least 4 characters");
		// TODO: Handle paragraph end markers (non-visible characters)
	});
	
	QUnit.test("GetParagraphs", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oPara1 = createParagraphWithText("First");
		const oPara2 = createParagraphWithText("Second");
		const oPara3 = createParagraphWithText("Third");
		oDocument.Push(oPara1);
		oDocument.Push(oPara2);
		oDocument.Push(oPara3);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const paragraphs = oSelection.GetParagraphs();
		
		assert.ok(paragraphs.length === 3, "Should return 3 paragraphs");
	});
	
	QUnit.test("GetHyperlinks", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		
		const oParagraph = createParagraphWithText("Visit ");
		oDocument.Push(oParagraph);
		oParagraph.AddHyperlink("https://onlyoffice.com", "Example Site");
		
		const oSelection = oDocument.GetSelection();
		let hyperlinks = oSelection.GetHyperlinks();
		assert.strictEqual(hyperlinks.length, 0, "Should return 0 hyperlinks");
		
		selectAllContent();
		hyperlinks = oSelection.GetHyperlinks();
		
		assert.strictEqual(hyperlinks.length, 1, "Should return 1 hyperlink");
		assert.strictEqual(hyperlinks[0].GetClassType(), "hyperlink", "Should be a hyperlink");
	});
	
	QUnit.test("GetMaths", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oParagraph = createParagraphWithText("No math here");
		oDocument.Push(oParagraph);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		let maths = oSelection.GetMaths();
		assert.strictEqual(maths.length, 0, "Should return 0 math formulas");
		
		oDocument.AddMathEquation('pi = 3', 'unicode');
		selectAllContent();
		maths = oSelection.GetMaths();
		assert.strictEqual(maths.length, 1, "Should return 1 math formula");
		assert.strictEqual(maths[0].GetClassType(), "math", "Should be a math object");
	});
	
	QUnit.test("GetComments", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oParagraph = createParagraphWithText("Text without comment");
		oDocument.Push(oParagraph);
		
		const oSelection = oDocument.GetSelection();
		
		selectAllContent();
		let comments = oSelection.GetComments();
		assert.strictEqual(comments.length, 0, "Should return 0 comments");
		
		oParagraph.AddComment("This is a comment");
		selectAllContent();
		comments = oSelection.GetComments();
		assert.strictEqual(comments.length, 1, "Should return 1 comment");
		assert.strictEqual(comments[0].GetClassType(), "comment", "Should be a comment object");
	});
	
	QUnit.test("GetText", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oPara1 = createParagraphWithText("First paragraph");
		const oPara2 = createParagraphWithText("Second paragraph");
		oDocument.Push(oPara1);
		oDocument.Push(oPara2);
		
		selectAllContent();
		const oSelection = oDocument.GetSelection();
		const text = oSelection.GetText();
		assert.strictEqual(text, "First paragraph\r\nSecond paragraph\r\n", "Selection text should match the combined paragraph texts");
	});
	
	QUnit.test("GetFields", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oParagraph = createParagraphWithText("Page field here: PAGE. Time field here: TIME.");
		oDocument.Push(oParagraph);
		
		const words = oParagraph.GetWords();
		
		let oRange1 = words[3].GetRange();
		oRange1.SetBold(true);
		oRange1.AddField("PAGE");
		
		let oRange2 = words[7].GetRange();
		oRange2.SetBold(true);
		oRange2.AddField('TIME \\@ "dddd, MMMM d, yyyy"');
		
		oParagraph.Select();
		const oSelection = oDocument.GetSelection();
		const fields = oSelection.GetFields();
		
		assert.strictEqual(fields.length, 2, "Should return 2 fields");
		assert.ok(fields[0].GetClassType() === "field" || fields[0].GetClassType() === "complexField",
			"First element should be a field");
		assert.ok(fields[1].GetClassType() === "field" || fields[1].GetClassType() === "complexField",
			"Second element should be a field");
	});
	
	QUnit.test("GetInlineShapes", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oParagraph = createParagraphWithText("Here is an inline image: ");
		oDocument.Push(oParagraph);
		
		const oImage1 = AscTest.Editor.CreateImage(
			'https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png',
			60 * 36000, 60 * 36000
		);
		oParagraph.AddDrawing(oImage1);
		oParagraph.AddText(' in the text.');
		
		const oSelection = oDocument.GetSelection();
		let inlineShapes = oSelection.GetInlineShapes();
		assert.strictEqual(inlineShapes.GetCount(), 0, "0 for empty selection");
		
		oParagraph.Select();
		inlineShapes = oSelection.GetInlineShapes();
		assert.strictEqual(inlineShapes.GetCount(), 1, "Image1 must be included");
		
		const oParagraph2 = createParagraphWithText("Another paragraph with images: ");
		oDocument.Push(oParagraph2);
		
		const oImage2 = AscTest.Editor.CreateImage(
			'https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png',
			30 * 36000, 30 * 36000
		);
		oParagraph2.AddDrawing(oImage2);
		oParagraph2.AddText(' and ');
		
		const oImage3 = AscTest.Editor.CreateImage(
			'https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png',
			20 * 36000, 20 * 36000
		);
		oParagraph2.AddDrawing(oImage3);
		
		selectAllContent();
		inlineShapes = oSelection.GetInlineShapes();
		assert.strictEqual(inlineShapes.GetCount(), 3, "Must return all 3 inline shapes");
	});
	
	QUnit.test("GetShapeRange", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		
		const oSelection = oDocument.GetSelection();
		let shapeRange = oSelection.GetShapeRange();
		assert.strictEqual(shapeRange.GetCount(), 0, "0 for empty selection");
		
		// Image
		let oParagraph = createParagraphWithText("Paragraph with floating image.");
		oDocument.Push(oParagraph);
		
		const oImage = AscTest.Editor.CreateImage(
			'https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png',
			60 * 36000, 60 * 36000
		);
		oImage.SetWrappingStyle('square');
		oParagraph.AddDrawing(oImage);
		
		oImage.Select();
		shapeRange = oSelection.GetShapeRange();
		assert.strictEqual(shapeRange.GetCount(), 1, "1 floating image");

		// Shape
		oParagraph = createParagraphWithText("Paragraph with floating shape.");
		oDocument.Push(oParagraph);
		
		const oShape = AscTest.Editor.CreateShape("rect", 60 * 36000, 35 * 36000, null, null);
		oShape.SetWrappingStyle('topAndBottom');
		oParagraph.AddDrawing(oShape);
		
		oShape.Select();
		shapeRange = oSelection.GetShapeRange();
		assert.strictEqual(shapeRange.GetCount(), 1, "1 floating shape");
		
		// Inline Image
		oParagraph = createParagraphWithText("Inline image: ");
		oDocument.Push(oParagraph);
		
		const oInlineImage = AscTest.Editor.CreateImage(
			'https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png',
			30 * 36000, 30 * 36000
		);
		oInlineImage.SetWrappingStyle('inline');
		oParagraph.AddDrawing(oInlineImage);
		
		oInlineImage.Select();
		shapeRange = oSelection.GetShapeRange();
		assert.strictEqual(shapeRange.GetCount(), 0, "Inline image should not be counted");
	});
	
	QUnit.test("GetTables, GetRows, GetCells", function (assert) {
		const oDocument = AscTest.Editor.GetDocument();
		const oSelection = oDocument.GetSelection();
		
		const oTable1 = AscTest.Editor.CreateTable(2, 2);
		const oTable2 = AscTest.Editor.CreateTable(3, 3);
		oDocument.Push(oTable1);
		oDocument.Push(createParagraphWithText("Between tables"));
		oDocument.Push(oTable2);
		
		let tables = oSelection.GetTables();
		assert.strictEqual(tables.length, 0, "Should return 0 tables");
		
		let rows = oSelection.GetRows();
		assert.strictEqual(rows.length, 0, "Should return 0 rows");
		
		let cells = oSelection.GetCells();
		assert.strictEqual(cells.length, 0, "Should return 0 cells");
		
		selectAllContent();
		tables = oSelection.GetTables();
		
		assert.strictEqual(tables.length, 2, "Must return 2 tables");
		assert.strictEqual(tables[0].GetClassType(), "table", "First should be a table");
		assert.strictEqual(tables[1].GetClassType(), "table", "Second should be a table");
		
		rows = oSelection.GetRows();
		assert.strictEqual(rows.length, 5, "Must return 5 rows total (2 + 3)");
		rows.forEach(function (row) {
			assert.strictEqual(row.GetClassType(), "tableRow", "Must be a table row");
		});
		
		cells = oSelection.GetCells();
		assert.strictEqual(cells.length, 13, "Should return 13 cells total");
		cells.forEach(function (cell) {
			assert.strictEqual(cell.GetClassType(), "tableCell", "Should be a table cell");
		});
	});
});
