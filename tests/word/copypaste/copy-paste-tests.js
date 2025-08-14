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
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"139"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Title"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Paragraph with bold and italic text."],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"numPr":{"ilvl":0,"numId":"495"},"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"165"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["List item 1"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["List item 2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		};

		assert.strictEqual(result, JSON.stringify(expected), "Complex HTML content should match expected JSON format");

		done();
	});

	QUnit.module("Word Copy Paste Tests");
});

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
