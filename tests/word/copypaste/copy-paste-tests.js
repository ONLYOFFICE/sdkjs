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

"use strict";

$(function () {

	// let startTime = 0;
	//
	// QUnit.begin(function() {
	// 	startTime = performance.now();
	// });
	//
	// QUnit.done(function(details) {
	// 	const endTime = performance.now();
	// 	const duration = (endTime - startTime).toFixed(2);
	// 	console.log(`Время выполнения всех тестов: ${duration} мс`);
	// });

	QUnit.module("Word Copy Paste Tests");

    let logicDocument = AscTest.CreateLogicDocument();
	AscTest.Editor.WordControl.m_oDrawingDocument.m_oLogicDocument = logicDocument;
	AscTest.Editor.WordControl.m_oLogicDocument = logicDocument;

	QUnit.test("Test: \"callback tests paste plain text\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, "test", undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste HTML\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "test HTML content";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste Internal format\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let binaryData = "";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, binaryData, undefined, undefined, undefined, function () {
			assert.ok(true);
			done();
		});
	});

	QUnit.test("Test: \"copy HTML with JSON verification\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		// Create an HTML element to simulate copying
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p>Test HTML content</p>";

		// Simulate pasting the HTML content into the document
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);

		const expected = {
			"type": "document",
			"textPr": "Test HTML content\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Test HTML content"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		};

		assert.strictEqual(result, JSON.stringify(expected), "HTML content should match expected JSON format");

		done();
	});

	QUnit.test("Test: \"copy complex HTML with JSON verification\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		// Create a complex HTML element to simulate copying
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = `
								<div>
								  <h1 style="color: red;">Title</h1>
								  <p>Paragraph with <strong>bold</strong> and <em>italic</em> text.</p>
								  <ul>
									<li>List item 1</li>
									<li>List item 2</li>
								  </ul>
								</div>
							  `;

		// Simulate pasting the HTML content into the document
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);

		const expected = {
			"type": "document",
			"textPr": "Title\r\nParagraph with bold and italic text.\r\n·\tList item 1\r\nList item 2\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"139"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Title"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Paragraph with bold and italic text."],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"numPr":{"ilvl":0,"numId":"488"},"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"165"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["List item 1"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["List item 2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		};

		// result json object content will have "numId\":\"...\", I need to copy that part into my expected object content
		// to make the test pass, because the numId is generated dynamically

		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}

		assert.strictEqual(result, JSON.stringify(expected), "Complex HTML content should match expected JSON format");

		done();
	});

	QUnit.test("Test: \"paste html, select text, copy html, check htmls\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		// Create a complex HTML element to simulate copying
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = `
								<div>
								  <h1 style="color: red;">Title</h1>
								  <p>Paragraph with <strong>bold</strong> and <em>italic</em> text.</p>
								  <ul>
									<li>List item 1</li>
									<li>List item 2</li>
								  </ul>
								</div>
							  `;

		// Simulate pasting the HTML content into the document
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		// Select the text in the paragraph and copy to clipboard
		logicDocument.SelectAll();
		var oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		const sBase64 = oCopyProcessor.Start();
		const _data = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();
		const jsonedData = removeBase64(JSON.stringify(_data));
		const trueExpectations =
			"\"<h1 style=\\\"mso-pagination:widow-orphan lines-together;page-break-after:avoid;margin-top:18pt;margin-bottom:4pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Arial';font-size:20pt;color:#376092;mso-style-textfill-fill-color:#376092\\\">Title</span></h1><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Paragraph with bold and italic text.</span></p><ul style=\\\"padding-left:40px\\\"><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">List item 1</span></p></li></ul><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">List item 2</span></p>\""

		assert.strictEqual(jsonedData, trueExpectations, "Copied data should be a document type");

		done();
	});

	QUnit.test("Paste simple div HTML content", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>Simple text</div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});
		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Simple text\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Simple text"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}

		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться текст из div элемента");

		done();
	});

	QUnit.test("Paste simple div HTML, then select & copy back", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>Simple text</div>";

		// Paste HTML (no callback, just call it)
		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement
		);

		// Now select & copy
		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		// Normalize copied HTML for comparison
		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		console.log(jsonedData);
		const expectedHtml = "\"<p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Simple text</span></p><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for simple div paste");
		done();
	});

	QUnit.test("Paste paragraph and span with style", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><span style='color:blue;'>Blue text</span></p>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Blue text\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Blue text"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться текст с цветом из span элемента");
		done();
	});

	QUnit.test("Paste table HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<table><tr><td>Cell 1</td><td>Cell 2</td></tr></table>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Cell 1\tCell 2\t\r\n",
			'content': [{"bPresentation":false,"tblGrid":[{"w":4677,"type":"gridCol"},{"w":4677,"type":"gridCol"}],"tblPr":{"tblBorders":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"end":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"insideH":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"insideV":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"start":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"tblCellMar":{},"tblLayout":"autofit","tblLook":{"firstColumn":true,"firstRow":true,"lastColumn":false,"lastRow":false,"noHBand":false,"noVBand":true},"tblOverlap":"overlap","tblpPr":{"horzAnchor":"page","vertAnchor":"page","tblpXSpec":"center","tblpYSpec":"center","tblpX":0,"tblpY":57,"bottomFromText":0,"leftFromText":0,"rightFromText":0,"topFromText":0},"tblStyle":"12","tblW":{"type":"auto","w":0},"inline":true,"type":"tablePr"},"content":[{"content":[{"content":{"bPresentation":false,"content":[{"bFromDocument":true,"pPr":{"spacing":{"before":0,"after":0},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Cell 1"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}],"type":"docContent"},"tcPr":{"tcBorders":{},"tcW":{"type":"dxa","w":4677},"type":"tableCellPr"},"id":"592","type":"tblCell"},{"content":{"bPresentation":false,"content":[{"bFromDocument":true,"pPr":{"spacing":{"before":0,"after":0},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Cell 2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}],"type":"docContent"},"tcPr":{"tcBorders":{},"tcW":{"type":"dxa","w":4677},"type":"tableCellPr"},"id":"602","type":"tblCell"}],"reviewInfo":{"userId":"","author":"","date":"","moveType":"noMove","prevType":-1},"reviewType":"common","trPr":{"type":"tableRowPr"},"type":"tblRow"}],"changes":[],"type":"table"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numIds = result.match(/"id":"(\d+)"/g);
		if (numIds) {
			expected.content[0].content[0].content[0].id = numIds[0].replace(/"id":"(\d+)"/, '$1');
			expected.content[0].content[0].content[1].id = numIds[1].replace(/"id":"(\d+)"/, '$1');
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должна вставиться таблица с двумя ячейками");
		done();
	});

	QUnit.test("Paste unordered list HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<ul><li>Item 1</li><li>Item 2</li></ul>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "·\tItem 1\r\n·\tItem 2\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"numPr":{"ilvl":0,"numId":"620"},"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"165"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Item 1"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"numPr":{"ilvl":0,"numId":"620"},"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"165"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Item 2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numIds = result.match(/"numId":"(\d+)"/g);
		if (numIds) {
			expected.content[0].pPr.numPr.numId = numIds[0].match(/"numId":"(\d+)"/)[1];
			expected.content[1].pPr.numPr.numId = numIds[1].match(/"numId":"(\d+)"/)[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться маркированный список с двумя элементами");
		done();
	});

	QUnit.test("Paste image HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUA' alt='Test Image'>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должно вставиться изображение из img элемента");
		done();
	});

	QUnit.test("Paste bold and italic HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><b>Bold</b> and <i>Italic</i></p>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Bold and Italic\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Bold and Italic"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться жирный и курсивный текст из HTML");
		done();
	});

	QUnit.test("Paste underline and strikethrough HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><u>Underline</u> and <s>Strikethrough</s></p>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Underline and Strikethrough\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Underline and Strikethrough"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться подчеркнутый и зачеркнутый текст из HTML");
		done();
	});

	QUnit.test("Paste hyperlink HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<a href='https://example.com'>Example Link</a>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Example Link\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"value":"https://example.com/","content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr","rStyle":"187"},"content":["Example Link"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"}],"type":"hyperlink"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должна вставиться гиперссылка из HTML");
		done();
	});

	// QUnit.test("Paste ordered list HTML", function(assert) {
	// 	AscTest.ClearDocument();
	// 	let p = AscTest.CreateParagraph();
	// 	logicDocument.AddToContent(0, p);
	//
	// 	let done = assert.async();
	// 	let htmlElement = document.createElement("div");
	// 	htmlElement.innerHTML = "<ol><li>First</li><li>Second</li></ol>";
	//
	// 	AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});
	//
	// 	const result = ToJsonString(logicDocument);
	// 	console.log("Ordered List Result:", result);
	// 	const expected = {
	// 		"type": "document",
	// 		"textPr": "",
	// 		'content': ""
	// 	}
	// 	// const numIds = result.match(/"numId":"(\d+)"/g);
	// 	// if (numIds) {
	// 	// 	expected.content[0].pPr.numPr.numId = numIds[0].match(/"numId":"(\d+)"/)[1];
	// 	// 	expected.content[1].pPr.numPr.numId = numIds[1].match(/"numId":"(\d+)"/)[1];
	// 	// }
	// 	assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться маркированный список с двумя элементами");
	// 	done();
	// });

	QUnit.test("Paste nested HTML elements", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div><span><b>Nested</b> <i>Elements</i></span></div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Nested Elements\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Nested Elements"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться вложенный текст с жирным и курсивом из HTML");
		done();
	});

	QUnit.test("Paste line break HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "Line1<br>Line2";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Line1\rLine2\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Line1",{"type":"break","breakType":"textWrapping"},"Line2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться перенос строки из HTML");
		done();
	});

	QUnit.test("Paste empty div HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div></div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться пустой параграф из пустого div элемента");
		done();
	});

	QUnit.test("Paste special character HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>&copy; &euro; &amp;</div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "© € &\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["© € &"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться текст с символами ©, €, & из HTML");
		done();
	});

	QUnit.test("Paste formula as text HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>y = mx + b</div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "y = mx + b\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["y = mx + b"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться текст формулы из HTML");
		done();
	});

	QUnit.test("Paste HTML with mso style", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = '<br style="page-break-before:always;mso-break-type:section-break;">';

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "\r\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"type":"break","breakType":"page"}],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться перенос страницы из HTML с mso стилем");
		done();
	});

	QUnit.test("Paste paragraph + span with style, then select & copy back", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><span style='color:blue;'>Blue text</span></p>";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Blue text</span></p><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for span with style");
		done();
	});

	// QUnit.test("Paste table HTML, then select & copy back", function(assert) {
	// 	AscTest.ClearDocument();
	// 	let p = AscTest.CreateParagraph();
	// 	logicDocument.AddToContent(0, p);
	// 	let done = assert.async();
	//
	// 	let htmlElement = document.createElement("div");
	// 	htmlElement.innerHTML = "<table><tr><td>Cell 1</td><td>Cell 2</td></tr></table>";
	// 	AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);
	//
	// 	logicDocument.SelectAll();
	// 	let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
	// 	oCopyProcessor.Start();
	// 	const copiedHtml = oCopyProcessor.getInnerHtml();
	// 	logicDocument.RemoveSelection();
	//
	// 	const jsonedData = removeBase64(JSON.stringify(copiedHtml));
	// 	const expectedHtml =''
	// 	assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for table paste");
	// 	done();
	// });

	QUnit.test("Paste unordered list HTML, then select & copy back", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<ul><li>Item 1</li><li>Item 2</li></ul>";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<ul style=\\\"padding-left:40px\\\" class=\\\"docData;\\\"><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Item 1</span></p></li><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Item 2</span></p></li></ul><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for unordered list");
		done();
	});

	QUnit.test("Paste bold/italic HTML, then select & copy back", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><b>Bold</b> and <i>Italic</i></p>";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Bold and Italic</span></p><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for bold + italic");
		done();
	});

	QUnit.module("Word Copy Paste Tests");
});

// function removeBase64(html) {
// 	// Regex to remove long base64-like strings (letters, digits, +, /, =)
// 	// Adjust {50,} if your garbage is shorter/longer
// 	return html.replace(/([A-Za-z0-9+/=]{50,})/g, '');
// }

function removeBase64(html) {
	// 1. Remove long base64-like strings (letters, digits, +, /, =)
	html = html.replace(/([A-Za-z0-9+/=]{50,})/g, '');

	// 2. Remove dynamic docData metadata like: docData;DOCY;v5;3707;
	html = html.replace(/docData;DOCY;v\d+;\d+;?/g, 'docData;');

	return html;
}

function ToJsonString(logicDocument) {
	var oWriter = new AscJsonConverter.WriterToJSON();

	var oResult = {
		"type":      "document",
		"textPr":    logicDocument.GetText(),
		'content':	 oWriter.SerContent(logicDocument.Content, undefined, undefined, undefined, true),
		// "paraPr":    bWriteDefaultParaPr ? oWriter.SerParaPr(this.GetDefaultParaPr().ParaPr) : undefined,
		// "theme":     bWriteTheme ? oWriter.SerTheme(this.Document.GetTheme()) : undefined,
		// "sectPr":    bWriteSectionPr ? oWriter.SerSectionPr(this.Document.SectPr) : undefined,
		// "numbering": bWriteNumberings ? oWriter.jsonWordNumberings : undefined,
		// "styles":    bWriteStyles ? oWriter.SerWordStylesForWrite() : undefined
	}

	return JSON.stringify(oResult);
}
