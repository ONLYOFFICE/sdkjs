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
	QUnit.module("Word Copy/Paste Tests");

	let logicDocument = AscTest.CreateLogicDocument();
	AscTest.Editor.WordControl.m_oDrawingDocument.m_oLogicDocument =
		logicDocument;
	AscTest.Editor.WordControl.m_oLogicDocument = logicDocument;

	// -------------------------------------------------------------------------
	// Helpers
	// -------------------------------------------------------------------------
	const charWidth = AscTest.CharWidth * AscTest.FontSize;
	const L_FIELD = 20 * charWidth;
	const R_FIELD = 30 * charWidth;
	const PAGE_W = 150 * charWidth;
	function initDocument() {
		AscTest.ClearDocument();
		logicDocument.AddToContent(0, AscTest.CreateParagraph());

		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(PAGE_W, 1000);
		sectPr.SetPageMargins(L_FIELD, 50, R_FIELD, 50);
	}
	const oldPrepeare_recursive = AscCommon.PasteProcessor.prototype._Prepeare_recursive;

	AscCommon.PasteProcessor.prototype._Prepeare_recursive = function () {};

	AscCommon.g_font_loader.CheckFontsNeedLoading = function () {};

	QUnit.test('Test: "callback tests paste plain text"', function (assert) {
		initDocument();

		let done = assert.async();
		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.Text,
			"test",
			undefined,
			undefined,
			undefined,
			function (success) {
				assert.ok(success);
				done();
			}
		);
	});

	QUnit.test('Test: "callback tests paste HTML"', function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "test HTML content";
		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
			undefined,
			undefined,
			undefined,
			function (success) {
				assert.ok(success);
				done();
			}
		);
	});

	QUnit.test(
		'Test: "callback tests paste Internal format"',
		function (assert) {
			initDocument();

			let done = assert.async();
			let binaryData = "";
			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.Internal,
				binaryData,
				undefined,
				undefined,
				undefined,
				function () {
					assert.ok(true);
					done();
				}
			);
		}
	);

	let defJsonBorders = {
		bottom: {
			color: { auto: false, r: 0, g: 0, b: 0 },
			sz: 4,
			space: 0,
			value: "none",
		},
		left: {
			color: { auto: false, r: 0, g: 0, b: 0 },
			sz: 4,
			space: 0,
			value: "none",
		},
		right: {
			color: { auto: false, r: 0, g: 0, b: 0 },
			sz: 4,
			space: 0,
			value: "none",
		},
		top: {
			color: { auto: false, r: 0, g: 0, b: 0 },
			sz: 4,
			space: 0,
			value: "none",
		},
	}
	
	QUnit.test('Test: "copy HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		// Create an HTML element to simulate copying
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p>Test HTML content</p>";

		// Simulate pasting the HTML content into the document
		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
			undefined,
			undefined,
			undefined,
			function () {}
		);

		const result = ToJsonString(logicDocument);

		const expected = {
			type: "document",
			textPr: "Test HTML content\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["Test HTML content"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy complex HTML with JSON verification"', function (assert) {
			initDocument();

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
			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.HtmlElement,
				htmlElement,
				undefined,
				undefined,
				undefined,
				function () {}
			);

			const result = ToJsonString(logicDocument);

			const expected = {
				type: "document",
				textPr: "Title\r\nParagraph with bold and italic text.\r\n·\tList item 1\r\nList item 2\r\n",
				content: [
					{
						bFromDocument: true,
						pPr: {
							pBdr: defJsonBorders,
							bFromDocument: true,
							type: "paraPr",
							pStyle: "139",
						},
						rPr: { bFromDocument: true, type: "textPr" },
						content: [
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: ["Title"],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "run",
							},
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: [],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "endRun",
							},
						],
						changes: [],
						type: "paragraph",
					},
					{
						bFromDocument: true,
						pPr: {
							pBdr: defJsonBorders,
							bFromDocument: true,
							type: "paraPr",
						},
						rPr: { bFromDocument: true, type: "textPr" },
						content: [
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: [
									"Paragraph with bold and italic text.",
								],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "run",
							},
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: [],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "endRun",
							},
						],
						changes: [],
						type: "paragraph",
					},
					{
						bFromDocument: true,
						pPr: {
							numPr: { ilvl: 0, numId: "488" },
							pBdr: defJsonBorders,
							bFromDocument: true,
							type: "paraPr",
							pStyle: "165",
						},
						rPr: { bFromDocument: true, type: "textPr" },
						content: [
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: ["List item 1"],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "run",
							},
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: [],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "endRun",
							},
						],
						changes: [],
						type: "paragraph",
					},
					{
						bFromDocument: true,
						pPr: { bFromDocument: true, type: "paraPr" },
						rPr: { bFromDocument: true, type: "textPr" },
						content: [
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: ["List item 2"],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "run",
							},
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: [],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "run",
							},
							{
								bFromDocument: true,
								rPr: { bFromDocument: true, type: "textPr" },
								content: [],
								footnotes: [],
								endnotes: [],
								reviewType: "common",
								type: "endRun",
							},
						],
						changes: [],
						type: "paragraph",
					},
				],
			};

			// result json object content will have "numId\":\"...\", I need to copy that part into my expected object content
			// to make the test pass, because the numId is generated dynamically

			const numId = result.match(/"numId":"(\d+)"/);
			if (numId) {
				expected.content[2].pPr.numPr.numId = numId[1];
			}

			assert.strictEqual(
				result,
				JSON.stringify(expected),
				"Complex HTML content should match expected JSON format"
			);

			done();
		}
	);

	QUnit.test('Test: "paste html, select text, copy html, check htmls for simple lists"', function (assert) {
			initDocument();

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
			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.HtmlElement,
				htmlElement,
				undefined,
				undefined,
				undefined,
				function () {}
			);

			// Select the text in the paragraph and copy to clipboard
			logicDocument.SelectAll();
			var oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
			const sBase64 = oCopyProcessor.Start();
			const _data = oCopyProcessor.getInnerHtml();
			logicDocument.RemoveSelection();
			const jsonedData = removeBase64(JSON.stringify(_data));
			const trueExpectations =
				'"<h1 style=\\"mso-pagination:widow-orphan lines-together;page-break-after:avoid;margin-top:18pt;margin-bottom:4pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\" class=\\"docData;\\"><span style=\\"font-family:\'Arial\';font-size:20pt;color:#376092;mso-style-textfill-fill-color:#376092\\">Title</span></h1><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Paragraph with bold and italic text.</span></p><ul style=\\"padding-left:40px\\"><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">List item 1</span></p></li></ul><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">List item 2</span></p>"';

			assert.strictEqual(
				jsonedData,
				trueExpectations,
				"Copied data should be a document type"
			);

			done();
		}
	);

	QUnit.test('Test: "paste html, select text, copy html, check htmls for marked lists"', function (assert) {
			initDocument();

			let done = assert.async();

			// Insert bulleted list only
			let htmlElement = document.createElement("div");
			htmlElement.innerHTML = `
					<ul>
						<li>Элемент 1</li>
						<li>Элемент 2</li>
						<li>Элемент 3</li>
					</ul>`;

			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.HtmlElement,
				htmlElement,
				undefined,
				undefined,
				undefined,
				function () {
					// Copy back
					logicDocument.SelectAll();
					var oCopyProcessor = new AscCommon.CopyProcessor(
						AscTest.Editor
					);
					oCopyProcessor.Start();
					const copiedHtml = oCopyProcessor.getInnerHtml();
					logicDocument.RemoveSelection();

					const jsonedData = removeBase64(JSON.stringify(copiedHtml));
					const expectedHtml =
						'"<ul style=\\"padding-left:40px\\" class=\\"docData;\\"><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Элемент 1</span></p></li><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Элемент 2</span></p></li><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Элемент 3</span></p></li></ul><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\">&nbsp;</p>"';
					assert.strictEqual(
						jsonedData,
						expectedHtml,
						"Should correctly copy unordered list"
					);
					done();
				}
			);
		}
	);

	QUnit.test(
		'Test: "paste html, select text, copy html, check htmls for numbered lists"',
		function (assert) {
			initDocument();

			let done = assert.async();

			// Insert ordered list only
			let htmlElement = document.createElement("div");
			htmlElement.innerHTML = `
					<ol>
					  <li>Элемент 1</li>
					  <li>Элемент 2</li>
					  <li>Элемент 3</li>
					</ol>
				  `;

			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.HtmlElement,
				htmlElement,
				undefined,
				undefined,
				undefined,
				function () {
					// Copy back
					logicDocument.SelectAll();
					var oCopyProcessor = new AscCommon.CopyProcessor(
						AscTest.Editor
					);
					oCopyProcessor.Start();
					const copiedHtml = oCopyProcessor.getInnerHtml();
					logicDocument.RemoveSelection();

					const jsonedData = removeBase64(JSON.stringify(copiedHtml));
					const expectedHtml = "\"<ol style=\\\"padding-left:40px\\\" class=\\\"docData;\\\"><li style=\\\"list-style-type: decimal\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Элемент 1</span></p></li><li style=\\\"list-style-type: decimal\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Элемент 2</span></p></li></ol><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Элемент 3</span></p>\"";
					assert.strictEqual(
						jsonedData,
						expectedHtml,
						"Should correctly copy numbered list"
					);
					done();
				}
			);
		}
	);

	QUnit.test('Test: "paste html, select text, copy html, check htmls for multi-level lists"', function (assert) {
			initDocument();

			let done = assert.async();

			// Insert multi-level bulleted list
			let htmlElement = document.createElement("div");
			htmlElement.innerHTML = `
		<ul>
		  <li>Первый уровень 1
			<ul>
			  <li>Второй уровень 1</li>
			  <li>Второй уровень 2</li>
			</ul>
		  </li>
		  <li>Первый уровень 2</li>
		</ul>
	  `;

			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.HtmlElement,
				htmlElement,
				undefined,
				undefined,
				undefined,
				function () {
					logicDocument.SelectAll();
					var oCopyProcessor = new AscCommon.CopyProcessor(
						AscTest.Editor
					);
					oCopyProcessor.Start();
					const copiedHtml = oCopyProcessor.getInnerHtml();
					logicDocument.RemoveSelection();

					const jsonedData = removeBase64(JSON.stringify(copiedHtml));
					const expectedHtml =
						'"<ul style=\\"padding-left:40px\\" class=\\"docData;\\"><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Первый уровень 1&nbsp;&nbsp;&nbsp; </span></p></li><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:71.43307086614173pt;text-indent:-17.999999999999996pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Второй уровень 1</span></p></li><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:71.43307086614173pt;text-indent:-17.999999999999996pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Второй уровень 2</span></p></li></ul><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Первый уровень 2</span></p>"';
					assert.strictEqual(
						jsonedData,
						expectedHtml,
						"Should correctly copy multilevel list"
					);
					done();
				}
			);
		}
	);

	QUnit.test("Paste simple div HTML content", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>Simple text</div>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);
		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "Simple text\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["Simple text"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};

		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste text from div element"
		);

		done();
	});

	QUnit.test("Paste simple div HTML, then select & copy back", function (assert) {
			initDocument();

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
			const expectedHtml =
				'"<p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\" class=\\"docData;\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Simple text</span></p><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\">&nbsp;</p>"';
			assert.strictEqual(
				jsonedData,
				expectedHtml,
				"Copied HTML should match for simple div paste"
			);
			done();
		}
	);

	QUnit.test("Paste paragraph and span with style", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<p><span style='color:blue;'>Blue text</span></p>";


		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "Blue text\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["Blue text"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste colored text from span element"
		);
		done();
	});

	QUnit.test("Paste table HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<table><tr><td>Cell 1</td><td>Cell 2</td></tr></table>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "Cell 1\tCell 2\t\r\n",
			content: [
				{
					bPresentation: false,
					tblGrid: [
						{ w: 4677, type: "gridCol" },
						{ w: 4677, type: "gridCol" },
					],
					tblPr: {
						tblBorders: {
							bottom: {
								color: { auto: false, r: 0, g: 0, b: 0 },
								sz: 4,
								space: 0,
								value: "none",
							},
							end: {
								color: { auto: false, r: 0, g: 0, b: 0 },
								sz: 4,
								space: 0,
								value: "none",
							},
							insideH: {
								color: { auto: false, r: 0, g: 0, b: 0 },
								sz: 4,
								space: 0,
								value: "none",
							},
							insideV: {
								color: { auto: false, r: 0, g: 0, b: 0 },
								sz: 4,
								space: 0,
								value: "none",
							},
							start: {
								color: { auto: false, r: 0, g: 0, b: 0 },
								sz: 4,
								space: 0,
								value: "none",
							},
							top: {
								color: { auto: false, r: 0, g: 0, b: 0 },
								sz: 4,
								space: 0,
								value: "none",
							},
						},
						tblCellMar: {},
						tblLayout: "autofit",
						tblLook: {
							firstColumn: true,
							firstRow: true,
							lastColumn: false,
							lastRow: false,
							noHBand: false,
							noVBand: true,
						},
						tblOverlap: "overlap",
						tblpPr: {
							horzAnchor: "page",
							vertAnchor: "page",
							tblpXSpec: "center",
							tblpYSpec: "center",
							tblpX: 0,
							tblpY: 57,
							bottomFromText: 0,
							leftFromText: 0,
							rightFromText: 0,
							topFromText: 0,
						},
						tblStyle: "12",
						tblW: { type: "auto", w: 0 },
						inline: true,
						type: "tablePr",
					},
					content: [
						{
							content: [
								{
									content: {
										bPresentation: false,
										content: [
											{
												bFromDocument: true,
												pPr: {
													spacing: {
														before: 0,
														after: 0,
													},
													bFromDocument: true,
													type: "paraPr",
												},
												rPr: {
													bFromDocument: true,
													type: "textPr",
												},
												content: [
													{
														bFromDocument: true,
														rPr: {
															bFromDocument: true,
															type: "textPr",
														},
														content: ["Cell 1"],
														footnotes: [],
														endnotes: [],
														reviewType: "common",
														type: "run",
													},
													{
														bFromDocument: true,
														rPr: {
															bFromDocument: true,
															type: "textPr",
														},
														content: [],
														footnotes: [],
														endnotes: [],
														reviewType: "common",
														type: "endRun",
													},
												],
												changes: [],
												type: "paragraph",
											},
										],
										type: "docContent",
									},
									tcPr: {
										tcBorders: {},
										tcW: { type: "dxa", w: 4677 },
										type: "tableCellPr",
									},
									id: "592",
									type: "tblCell",
								},
								{
									content: {
										bPresentation: false,
										content: [
											{
												bFromDocument: true,
												pPr: {
													spacing: {
														before: 0,
														after: 0,
													},
													bFromDocument: true,
													type: "paraPr",
												},
												rPr: {
													bFromDocument: true,
													type: "textPr",
												},
												content: [
													{
														bFromDocument: true,
														rPr: {
															bFromDocument: true,
															type: "textPr",
														},
														content: ["Cell 2"],
														footnotes: [],
														endnotes: [],
														reviewType: "common",
														type: "run",
													},
													{
														bFromDocument: true,
														rPr: {
															bFromDocument: true,
															type: "textPr",
														},
														content: [],
														footnotes: [],
														endnotes: [],
														reviewType: "common",
														type: "endRun",
													},
												],
												changes: [],
												type: "paragraph",
											},
										],
										type: "docContent",
									},
									tcPr: {
										tcBorders: {},
										tcW: { type: "dxa", w: 4677 },
										type: "tableCellPr",
									},
									id: "602",
									type: "tblCell",
								},
							],
							reviewInfo: {
								userId: "",
								author: "",
								date: "",
								moveType: "noMove",
								prevType: -1,
							},
							reviewType: "common",
							trPr: { type: "tableRowPr" },
							type: "tblRow",
						},
					],
					changes: [],
					type: "table",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numIds = result.match(/"id":"(\d+)"/g);
		if (numIds) {
			expected.content[0].content[0].content[0].id = numIds[0].replace(
				/"id":"(\d+)"/,
				"$1"
			);
			expected.content[0].content[0].content[1].id = numIds[1].replace(
				/"id":"(\d+)"/,
				"$1"
			);
		}
		const wMatches = result.match(/"w":(\d+)/g);
		if (wMatches && wMatches.length >= 5) {
			const wValues = wMatches.map((m) => m.replace(/"w":/, ""));

			// patch the 4 locations in expected
			expected.content[0].tblGrid[0].w = Number(wValues[0]);
			expected.content[0].tblGrid[1].w = Number(wValues[1]);
			expected.content[0].content[0].content[0].tcPr.tcW.w = Number(
				wValues[3]
			);
			expected.content[0].content[0].content[1].tcPr.tcW.w = Number(
				wValues[4]
			);
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste table with two cells"
		);
		done();
	});

	QUnit.test("Paste unordered list HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<ul><li>Item 1</li><li>Item 2</li></ul>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "·\tItem 1\r\n·\tItem 2\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						numPr: { ilvl: 0, numId: "620" },
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
						pStyle: "165",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["Item 1"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: {
						numPr: { ilvl: 0, numId: "620" },
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
						pStyle: "165",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["Item 2"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numIds = result.match(/"numId":"(\d+)"/g);
		if (numIds) {
			expected.content[0].pPr.numPr.numId =
				numIds[0].match(/"numId":"(\d+)"/)[1];
			expected.content[1].pPr.numPr.numId =
				numIds[1].match(/"numId":"(\d+)"/)[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste unordered list with two items"
		);
		done();
	});

	QUnit.test("Paste image HTML", function (assert) {
		// add prepare recursive to tests
		AscCommon.PasteProcessor.prototype._Prepeare_recursive = oldPrepeare_recursive;
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUA' alt='Test Image'>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste image from img element"
		);

		// remove prepeare recursive from tests
		AscCommon.PasteProcessor.prototype._Prepeare_recursive = function () {};
		done();
	});

	QUnit.test("Paste bold and italic HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><b>Bold</b> and <i>Italic</i></p>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "Bold and Italic\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["Bold and Italic"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste bold and italic text from HTML"
		);
		done();
	});

	QUnit.test("Paste underline and strikethrough HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<p><u>Underline</u> and <s>Strikethrough</s></p>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "Underline and Strikethrough\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["Underline and Strikethrough"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste underlined and strikethrough text from HTML"
		);
		done();
	});

	QUnit.test("Paste hyperlink HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<a href='https://example.com'>Example Link</a>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "Example Link\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							value: "https://example.com/",
							content: [
								{
									bFromDocument: true,
									rPr: {
										bFromDocument: true,
										type: "textPr",
										rStyle: "187",
									},
									content: ["Example Link"],
									footnotes: [],
									endnotes: [],
									reviewType: "common",
									type: "run",
								},
							],
							type: "hyperlink",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste hyperlink from HTML"
		);
		done();
	});

	QUnit.test("Paste nested HTML elements", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<div><span><b>Nested</b> <i>Elements</i></span></div>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "Nested Elements\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["Nested Elements"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste nested text with bold and italic from HTML"
		);
		done();
	});

	QUnit.test("Paste line break HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "Line1<br>Line2";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "Line1\rLine2\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [
								"Line1",
								{ type: "break", breakType: "textWrapping" },
								"Line2",
							],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste line break from HTML"
		);
		done();
	});

	QUnit.test("Paste empty div HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div></div>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste empty paragraph from empty div element"
		);
		done();
	});

	QUnit.test("Paste special character HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>&copy; &euro; &amp;</div>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "© € &\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["© € &"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste text with symbols ©, €, & from HTML"
		);
		done();
	});

	QUnit.test("Paste formula as text HTML", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>y = mx + b</div>";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "y = mx + b\r\n\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: {
						pBdr: defJsonBorders,
						bFromDocument: true,
						type: "paraPr",
					},
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: ["y = mx + b"],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste formula text from HTML"
		);
		done();
	});

	QUnit.test("Paste HTML with mso style", function (assert) {
		initDocument();

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			'<br style="page-break-before:always;mso-break-type:section-break;">';

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
		);

		const result = ToJsonString(logicDocument);
		const expected = {
			type: "document",
			textPr: "\r\r\n",
			content: [
				{
					bFromDocument: true,
					pPr: { bFromDocument: true, type: "paraPr" },
					rPr: { bFromDocument: true, type: "textPr" },
					content: [
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [{ type: "break", breakType: "page" }],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "run",
						},
						{
							bFromDocument: true,
							rPr: { bFromDocument: true, type: "textPr" },
							content: [],
							footnotes: [],
							endnotes: [],
							reviewType: "common",
							type: "endRun",
						},
					],
					changes: [],
					type: "paragraph",
				},
			],
		};
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"Should paste page break from HTML with mso style"
		);
		done();
	});

	QUnit.test("Paste paragraph + span with style, then select & copy back", function (assert) {
			initDocument();
			let done = assert.async();

			let htmlElement = document.createElement("div");
			htmlElement.innerHTML =
				"<p><span style='color:blue;'>Blue text</span></p>";
			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.HtmlElement,
				htmlElement
			);

			logicDocument.SelectAll();
			let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
			oCopyProcessor.Start();
			const copiedHtml = oCopyProcessor.getInnerHtml();
			logicDocument.RemoveSelection();

			const jsonedData = removeBase64(JSON.stringify(copiedHtml));
			const expectedHtml =
				'"<p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\" class=\\"docData;\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Blue text</span></p><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\">&nbsp;</p>"';
			assert.strictEqual(
				jsonedData,
				expectedHtml,
				"Copied HTML should match for span with style"
			);
			done();
		}
	);

	QUnit.test("Paste unordered list HTML, then select & copy back", function (assert) {
			initDocument();
			let done = assert.async();

			let htmlElement = document.createElement("div");
			htmlElement.innerHTML = "<ul><li>Item 1</li><li>Item 2</li></ul>";
			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.HtmlElement,
				htmlElement
			);

			logicDocument.SelectAll();
			let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
			oCopyProcessor.Start();
			const copiedHtml = oCopyProcessor.getInnerHtml();
			logicDocument.RemoveSelection();

			const jsonedData = removeBase64(JSON.stringify(copiedHtml));
			const expectedHtml =
				'"<ul style=\\"padding-left:40px\\" class=\\"docData;\\"><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Item 1</span></p></li><li style=\\"list-style-type: disc\\"><p style=\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Item 2</span></p></li></ul><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\">&nbsp;</p>"';
			assert.strictEqual(
				jsonedData,
				expectedHtml,
				"Copied HTML should match for unordered list"
			);
			done();
		}
	);

	QUnit.test("Paste bold/italic HTML, then select & copy back", function (assert) {
			initDocument();
			let done = assert.async();

			let htmlElement = document.createElement("div");
			htmlElement.innerHTML = "<p><b>Bold</b> and <i>Italic</i></p>";
			AscTest.Editor.asc_PasteData(
				AscCommon.c_oAscClipboardDataFormat.HtmlElement,
				htmlElement
			);

			logicDocument.SelectAll();
			let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
			oCopyProcessor.Start();
			const copiedHtml = oCopyProcessor.getInnerHtml();
			logicDocument.RemoveSelection();

			const jsonedData = removeBase64(JSON.stringify(copiedHtml));
			const expectedHtml =
				'"<p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\" class=\\"docData;\\"><span style=\\"font-family:\'Times New Roman\';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\">Bold and Italic</span></p><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\">&nbsp;</p>"';
			assert.strictEqual(
				jsonedData,
				expectedHtml,
				"Copied HTML should match for bold + italic"
			);
			done();
		}
	);

	QUnit.test("Paste sum formula from excel to word", function (assert) {
		initDocument(logicDocument);
		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = `<head>
			<meta http-equiv=Content-Type content="text/html; charset=utf-8">
			<meta name=ProgId content=Excel.Sheet>
			<meta name=Generator content="Microsoft Excel 15">
			<link id=Main-File rel=Main-File
			href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip.htm">
			<link rel=File-List
			href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_filelist.xml">
			<style>
			<!--table
			\t{mso-displayed-decimal-separator:"\\,";
			\tmso-displayed-thousand-separator:" ";}
			@page
			\t{margin:.75in .7in .75in .7in;
			\tmso-header-margin:.3in;
			\tmso-footer-margin:.3in;}
			tr
			\t{mso-height-source:auto;}
			col
			\t{mso-width-source:auto;}
			br
			\t{mso-data-placement:same-cell;}
			td
			\t{padding-top:1px;
			\tpadding-right:1px;
			\tpadding-left:1px;
			\tmso-ignore:padding;
			\tcolor:black;
			\tfont-size:11.0pt;
			\tfont-weight:400;
			\tfont-style:normal;
			\ttext-decoration:none;
			\tfont-family:Calibri, sans-serif;
			\tmso-font-charset:0;
			\tmso-number-format:General;
			\ttext-align:general;
			\tvertical-align:bottom;
			\tborder:none;
			\tmso-background-source:auto;
			\tmso-pattern:auto;
			\tmso-protection:locked visible;
			\twhite-space:nowrap;
			\tmso-rotate:0;}
			-->
			</style>
			</head>

			<body link="#0563C1" vlink="#954F72">

			<table border=0 cellpadding=0 cellspacing=0 width=256 style='border-collapse:
			 collapse;width:192pt'>
			<!--StartFragment-->
			 <col width=64 span=4 style='width:48pt'>
			 <tr height=20 style='height:15.0pt'>
			  <td height=20 width=64 style='height:15.0pt;width:48pt'></td>
			  <td width=64 style='width:48pt'></td>
			  <td width=64 style='width:48pt'></td>
			  <td width=64 style='width:48pt'></td>
			 </tr>
			 <tr height=20 style='height:15.0pt'>
			  <td height=20 align=right style='height:15.0pt'>1</td>
			  <td align=right>2</td>
			  <td align=right>3</td>
			  <td align=right>6</td>
			 </tr>
			 <!--EndFragment-->
			</table>
			</body>`;

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement
		);

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<table cellspacing=\\\"0\\\" border=\\\"0\\\" cellpadding=\\\"0\\\" style=\\\"margin-left:0pt;mso-padding-alt:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-insidev:none;mso-border-insideh:none;\\\" class=\\\"docData;\\\"><tr style=\\\"height:15pt;\\\"><td width=\\\"472\\\" style=\\\"width:354.33070866141736pt;padding:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\">&nbsp;</p></td><td width=\\\"472\\\" style=\\\"width:354.33070866141736pt;padding:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\">&nbsp;</p></td><td width=\\\"472\\\" style=\\\"width:354.33070866141736pt;padding:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\">&nbsp;</p></td><td width=\\\"472\\\" style=\\\"width:354.33070866141736pt;padding:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\">&nbsp;</p></td></tr><tr style=\\\"height:15pt;\\\"><td width=\\\"472\\\" style=\\\"width:354.33070866141736pt;padding:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"text-align:right;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">1</span></p></td><td width=\\\"472\\\" style=\\\"width:354.33070866141736pt;padding:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"text-align:right;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">2</span></p></td><td width=\\\"472\\\" style=\\\"width:354.33070866141736pt;padding:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"text-align:right;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">3</span></p></td><td width=\\\"472\\\" style=\\\"width:354.33070866141736pt;padding:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"text-align:right;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">6</span></p></td></tr></table><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(
			jsonedData,
			expectedHtml,
			"Copied HTML should match for bold + italic"
		);
		done();
	});

	// QUnit.test("Paste Newton's binom formula from word", function (assert) {
	// 	initDocument(logicDocument);
	// 	let done = assert.async();
	//
	// 	// htmlElement.setAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
	// 	// htmlElement.setAttribute("xmlns:w", "urn:schemas-microsoft-com:office:word");
	// 	// htmlElement.setAttribute('xmlns:m', "http://schemas.microsoft.com/office/2004/12/omml");
	// 	// htmlElement.setAttribute("xmlns", "http://www.w3.org/TR/REC-html40");
	//
	// 	const htmlElement = document.createElement("div");
	// 	htmlElement.innerHTML = `
	// 			<head>
	// 		<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	// 		<meta name=ProgId content=Word.Document>
	// 		<meta name=Generator content="Microsoft Word 15">
	// 		<meta name=Originator content="Microsoft Word 15">
	// 		<link rel=File-List
	// 		href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_filelist.xml">
	// 		<link rel=themeData
	// 		href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_themedata.thmx">
	// 		<link rel=colorSchemeMapping
	// 		href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_colorschememapping.xml">
	//
	// 		<style>
	// 		<!--
	// 		 /* Font Definitions */
	// 		 @font-face
	// 		\t{font-family:"Cambria Math";
	// 		\tpanose-1:2 4 5 3 5 4 6 3 2 4;
	// 		\tmso-font-charset:204;
	// 		\tmso-generic-font-family:roman;
	// 		\tmso-font-pitch:variable;
	// 		\tmso-font-signature:-536869121 1107305727 33554432 0 415 0;}
	// 		@font-face
	// 		\t{font-family:Aptos;
	// 		\tmso-font-charset:0;
	// 		\tmso-generic-font-family:swiss;
	// 		\tmso-font-pitch:variable;
	// 		\tmso-font-signature:536871559 3 0 0 415 0;}
	// 		 /* Style Definitions */
	// 		 p.MsoNormal, li.MsoNormal, div.MsoNormal
	// 		\t{mso-style-unhide:no;
	// 		\tmso-style-qformat:yes;
	// 		\tmso-style-parent:"";
	// 		\tmargin-top:0cm;
	// 		\tmargin-right:0cm;
	// 		\tmargin-bottom:8.0pt;
	// 		\tmargin-left:0cm;
	// 		\tline-height:115%;
	// 		\tmso-pagination:widow-orphan;
	// 		\tfont-size:12.0pt;
	// 		\tfont-family:"Aptos",sans-serif;
	// 		\tmso-ascii-font-family:Aptos;
	// 		\tmso-ascii-theme-font:minor-latin;
	// 		\tmso-fareast-font-family:Aptos;
	// 		\tmso-fareast-theme-font:minor-latin;
	// 		\tmso-hansi-font-family:Aptos;
	// 		\tmso-hansi-theme-font:minor-latin;
	// 		\tmso-bidi-font-family:"Times New Roman";
	// 		\tmso-bidi-theme-font:minor-bidi;
	// 		\tmso-font-kerning:1.0pt;
	// 		\tmso-ligatures:standardcontextual;
	// 		\tmso-fareast-language:EN-US;}
	// 		.MsoChpDefault
	// 		\t{mso-style-type:export-only;
	// 		\tmso-default-props:yes;
	// 		\tfont-family:"Aptos",sans-serif;
	// 		\tmso-ascii-font-family:Aptos;
	// 		\tmso-ascii-theme-font:minor-latin;
	// 		\tmso-fareast-font-family:Aptos;
	// 		\tmso-fareast-theme-font:minor-latin;
	// 		\tmso-hansi-font-family:Aptos;
	// 		\tmso-hansi-theme-font:minor-latin;
	// 		\tmso-bidi-font-family:"Times New Roman";
	// 		\tmso-bidi-theme-font:minor-bidi;
	// 		\tmso-fareast-language:EN-US;}
	// 		.MsoPapDefault
	// 		\t{mso-style-type:export-only;
	// 		\tmargin-bottom:8.0pt;
	// 		\tline-height:115%;}
	// 		@page WordSection1
	// 		\t{size:595.3pt 841.9pt;
	// 		\tmargin:2.0cm 42.5pt 2.0cm 3.0cm;
	// 		\tmso-header-margin:35.4pt;
	// 		\tmso-footer-margin:35.4pt;
	// 		\tmso-paper-source:0;}
	// 		div.WordSection1
	// 		\t{page:WordSection1;}
	// 		-->
	// 		</style>
	// 		</head>
	//
	// 		<body lang=RU style='tab-interval:35.4pt;word-wrap:break-word'>
	// 		<!--StartFragment-->
	//
	// 		<p class=MsoNormal><!--[if gte msEquation 12]><m:oMathPara><m:oMath><m:sSup><m:sSupPr><span
	// 			style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
	// 			mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:sSupPr><m:e><m:d><m:dPr><span
	// 			  style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
	// 			  mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:dPr><m:e><i
	// 			  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif'><m:r>x</m:r><m:r>+</m:r><m:r>a</m:r></span></i></m:e></m:d></m:e><m:sup><i
	// 			style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif'><m:r>n</m:r></span></i></m:sup></m:sSup><i
	// 		  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 		  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>=</m:r></span></i><m:nary><m:naryPr><m:chr
	// 			 m:val="в€‘"/><m:grow m:val="on"/><span style='font-family:"Cambria Math",serif;
	// 			mso-ascii-font-family:"Cambria Math";mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:naryPr><m:sub><i
	// 			style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 			mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>k</m:r><m:r>=0</m:r></span></i></m:sub><m:sup><i
	// 			style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 			mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>n</m:r></span></i></m:sup><m:e><m:d><m:dPr><span
	// 			  style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
	// 			  mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:dPr><m:e><m:f><m:fPr><m:type
	// 				 m:val="noBar"/><span style='font-family:"Cambria Math",serif;
	// 				mso-ascii-font-family:"Cambria Math";mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:fPr><m:num><i
	// 				style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 				mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>n</m:r></span></i></m:num><m:den><i
	// 				style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 				mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>k</m:r></span></i></m:den></m:f></m:e></m:d><m:sSup><m:sSupPr><span
	// 			  style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
	// 			  mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:sSupPr><m:e><i
	// 			  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 			  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>x</m:r></span></i></m:e><m:sup><i
	// 			  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 			  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>k</m:r></span></i></m:sup></m:sSup><m:sSup><m:sSupPr><span
	// 			  style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
	// 			  mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:sSupPr><m:e><i
	// 			  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 			  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>a</m:r></span></i></m:e><m:sup><i
	// 			  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
	// 			  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>n</m:r><m:r>-</m:r><m:r>k</m:r></span></i></m:sup></m:sSup></m:e></m:nary></m:oMath></m:oMathPara><![endif]--><![if !msEquation]><span
	// 		style='font-size:12.0pt;line-height:115%;font-family:"Aptos",sans-serif;
	// 		mso-ascii-theme-font:minor-latin;mso-fareast-font-family:Aptos;mso-fareast-theme-font:
	// 		minor-latin;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:"Times New Roman";
	// 		mso-bidi-theme-font:minor-bidi;mso-ansi-language:RU;mso-fareast-language:EN-US;
	// 		mso-bidi-language:AR-SA'><v:shapetype id="_x0000_t75" coordsize="21600,21600"
	// 		 o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f"
	// 		 stroked="f">
	// 		 <v:stroke joinstyle="miter"/>
	// 		 <v:formulas>
	// 		  <v:f eqn="if lineDrawn pixelLineWidth 0"/>
	// 		  <v:f eqn="sum @0 1 0"/>
	// 		  <v:f eqn="sum 0 0 @1"/>
	// 		  <v:f eqn="prod @2 1 2"/>
	// 		  <v:f eqn="prod @3 21600 pixelWidth"/>
	// 		  <v:f eqn="prod @3 21600 pixelHeight"/>
	// 		  <v:f eqn="sum @0 0 1"/>
	// 		  <v:f eqn="prod @6 1 2"/>
	// 		  <v:f eqn="prod @7 21600 pixelWidth"/>
	// 		  <v:f eqn="sum @8 21600 0"/>
	// 		  <v:f eqn="prod @7 21600 pixelHeight"/>
	// 		  <v:f eqn="sum @10 21600 0"/>
	// 		 </v:formulas>
	// 		 <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
	// 		 <o:lock v:ext="edit" aspectratio="t"/>
	// 		</v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" style='width:136.5pt;
	// 		 height:41.25pt'>
	// 		 <v:imagedata src="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_image001.png"
	// 		  o:title="" chromakey="white"/>
	// 		</v:shape></span><![endif]><o:p></o:p></p>
	//
	// 		<!--EndFragment-->
	// 		</body>
	//
	// 	`;
	//
	// 	AscTest.Editor.asc_PasteData(
	// 		AscCommon.c_oAscClipboardDataFormat.HtmlElement,
	// 		htmlElement
	// 	);
	//
	// 	logicDocument.SelectAll();
	// 	let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
	// 	oCopyProcessor.Start();
	// 	const copiedHtml = oCopyProcessor.getInnerHtml();
	// 	logicDocument.RemoveSelection();
	//
	// 	const jsonedData = removeBase64(JSON.stringify(copiedHtml));
	// 	const expectedHtml =
	// 		'"<p style=\\"line-height:13.8pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\" class=\\"docData;\\"><span style=\\"font-family:\'Times New Roman\';font-size:12pt;color:#000000;mso-style-textfill-fill-color:#000000\\">&nbsp; </span></p><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\">&nbsp;</p>"';
	// 	assert.strictEqual(
	// 		jsonedData,
	// 		expectedHtml,
	// 		"Copied HTML should match for Newton's binom formula"
	// 	);
	// 	done();
	// });

	// QUnit.test("Paste footnote formula from word", function (assert) {
	// 	initDocument(logicDocument);
	// 	let done = assert.async();
	//
	// 	// htmlElement.setAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
	// 	// htmlElement.setAttribute("xmlns:w", "urn:schemas-microsoft-com:office:word");
	// 	// htmlElement.setAttribute('xmlns:m', "http://schemas.microsoft.com/office/2004/12/omml");
	// 	// htmlElement.setAttribute("xmlns", "http://www.w3.org/TR/REC-html40");
	//
	// 	const htmlElement = document.createElement("div");
	// 	htmlElement.innerHTML = `
	// 			<html xmlns:o="urn:schemas-microsoft-com:office:office"
	// 	xmlns:w="urn:schemas-microsoft-com:office:word"
	// 	xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
	// 	xmlns="http://www.w3.org/TR/REC-html40">
	//
	// 	<head>
	// 	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	// 	<meta name=ProgId content=Word.Document>
	// 	<meta name=Generator content="Microsoft Word 15">
	// 	<meta name=Originator content="Microsoft Word 15">
	// 	<link rel=File-List
	// 	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_filelist.xml">
	// 	<link rel=themeData
	// 	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_themedata.thmx">
	// 	<link rel=colorSchemeMapping
	// 	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_colorschememapping.xml">
	// 	<style>
	// 	<!--
	// 	 /* Font Definitions */
	// 	 @font-face
	// 	\t{font-family:"Cambria Math";
	// 	\tpanose-1:2 4 5 3 5 4 6 3 2 4;
	// 	\tmso-font-charset:204;
	// 	\tmso-generic-font-family:roman;
	// 	\tmso-font-pitch:variable;
	// 	\tmso-font-signature:-536869121 1107305727 33554432 0 415 0;}
	// 	@font-face
	// 	\t{font-family:Aptos;
	// 	\tmso-font-charset:0;
	// 	\tmso-generic-font-family:swiss;
	// 	\tmso-font-pitch:variable;
	// 	\tmso-font-signature:536871559 3 0 0 415 0;}
	// 	 /* Style Definitions */
	// 	 p.MsoNormal, li.MsoNormal, div.MsoNormal
	// 	\t{mso-style-unhide:no;
	// 	\tmso-style-qformat:yes;
	// 	\tmso-style-parent:"";
	// 	\tmargin-top:0cm;
	// 	\tmargin-right:0cm;
	// 	\tmargin-bottom:8.0pt;
	// 	\tmargin-left:0cm;
	// 	\tline-height:115%;
	// 	\tmso-pagination:widow-orphan;
	// 	\tfont-size:12.0pt;
	// 	\tfont-family:"Aptos",sans-serif;
	// 	\tmso-ascii-font-family:Aptos;
	// 	\tmso-ascii-theme-font:minor-latin;
	// 	\tmso-fareast-font-family:Aptos;
	// 	\tmso-fareast-theme-font:minor-latin;
	// 	\tmso-hansi-font-family:Aptos;
	// 	\tmso-hansi-theme-font:minor-latin;
	// 	\tmso-bidi-font-family:"Times New Roman";
	// 	\tmso-bidi-theme-font:minor-bidi;
	// 	\tmso-font-kerning:1.0pt;
	// 	\tmso-ligatures:standardcontextual;
	// 	\tmso-fareast-language:EN-US;}
	// 	span.MsoEndnoteReference
	// 	\t{mso-style-noshow:yes;
	// 	\tmso-style-priority:99;
	// 	\tvertical-align:super;}
	// 	.MsoChpDefault
	// 	\t{mso-style-type:export-only;
	// 	\tmso-default-props:yes;
	// 	\tfont-family:"Aptos",sans-serif;
	// 	\tmso-ascii-font-family:Aptos;
	// 	\tmso-ascii-theme-font:minor-latin;
	// 	\tmso-fareast-font-family:Aptos;
	// 	\tmso-fareast-theme-font:minor-latin;
	// 	\tmso-hansi-font-family:Aptos;
	// 	\tmso-hansi-theme-font:minor-latin;
	// 	\tmso-bidi-font-family:"Times New Roman";
	// 	\tmso-bidi-theme-font:minor-bidi;
	// 	\tmso-fareast-language:EN-US;}
	// 	.MsoPapDefault
	// 	\t{mso-style-type:export-only;
	// 	\tmargin-bottom:8.0pt;
	// 	\tline-height:115%;}
	// 	@page WordSection1
	// 	\t{size:612.0pt 792.0pt;
	// 	\tmargin:2.0cm 42.5pt 2.0cm 3.0cm;
	// 	\tmso-header-margin:36.0pt;
	// 	\tmso-footer-margin:36.0pt;
	// 	\tmso-paper-source:0;}
	// 	div.WordSection1
	// 	\t{page:WordSection1;}
	// 	-->
	// 	</style>
	// 	</head>
	//
	// 	<body lang=RU style='tab-interval:35.4pt;word-wrap:break-word'>
	// 	<!--StartFragment--><span class=MsoEndnoteReference><span
	// 	style='font-size:12.0pt;line-height:115%;font-family:"Aptos",sans-serif;
	// 	mso-ascii-theme-font:minor-latin;mso-fareast-font-family:Aptos;mso-fareast-theme-font:
	// 	minor-latin;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:"Times New Roman";
	// 	mso-bidi-theme-font:minor-bidi;mso-ansi-language:RU;mso-fareast-language:EN-US;
	// 	mso-bidi-language:AR-SA'><span style='mso-special-character:footnote'><![if !supportFootnotes]><span
	// 	class=MsoEndnoteReference><span style='font-size:12.0pt;line-height:115%;
	// 	font-family:"Aptos",sans-serif;mso-ascii-theme-font:minor-latin;mso-fareast-font-family:
	// 	Aptos;mso-fareast-theme-font:minor-latin;mso-hansi-theme-font:minor-latin;
	// 	mso-bidi-font-family:"Times New Roman";mso-bidi-theme-font:minor-bidi;
	// 	mso-ansi-language:RU;mso-fareast-language:EN-US;mso-bidi-language:AR-SA'>[1]</span></span><![endif]></span></span></span><span
	// 	style='font-size:12.0pt;line-height:115%;font-family:"Aptos",sans-serif;
	// 	mso-ascii-theme-font:minor-latin;mso-fareast-font-family:Aptos;mso-fareast-theme-font:
	// 	minor-latin;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:"Times New Roman";
	// 	mso-bidi-theme-font:minor-bidi;mso-ansi-language:EN-US;mso-fareast-language:
	// 	EN-US;mso-bidi-language:AR-SA'> <span lang=EN-US>Mittal, S., &amp; Vetter, J.
	// 	S. (2015). A survey of CPU-GPU heterogeneous computing techniques.&nbsp;</span></span><i><span
	// 	style='font-size:12.0pt;line-height:115%;font-family:"Aptos",sans-serif;
	// 	mso-ascii-theme-font:minor-latin;mso-fareast-font-family:Aptos;mso-fareast-theme-font:
	// 	minor-latin;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:"Times New Roman";
	// 	mso-bidi-theme-font:minor-bidi;mso-ansi-language:RU;mso-fareast-language:EN-US;
	// 	mso-bidi-language:AR-SA'>ACM Computing Surveys (CSUR)</span></i><span
	// 	style='font-size:12.0pt;line-height:115%;font-family:"Aptos",sans-serif;
	// 	mso-ascii-theme-font:minor-latin;mso-fareast-font-family:Aptos;mso-fareast-theme-font:
	// 	minor-latin;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:"Times New Roman";
	// 	mso-bidi-theme-font:minor-bidi;mso-ansi-language:RU;mso-fareast-language:EN-US;
	// 	mso-bidi-language:AR-SA'>,&nbsp;<i>47</i>(4), 1-35.</span><!--EndFragment-->
	// 	</body>
	//
	// 	</html>`;
	//
	// 	AscTest.Editor.asc_PasteData(
	// 		AscCommon.c_oAscClipboardDataFormat.HtmlElement,
	// 		htmlElement
	// 	);
	//
	// 	logicDocument.SelectAll();
	// 	let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
	// 	oCopyProcessor.Start();
	// 	const copiedHtml = oCopyProcessor.getInnerHtml();
	// 	logicDocument.RemoveSelection();
	//
	// 	const jsonedData = removeBase64(JSON.stringify(copiedHtml));
	// 	const expectedHtml = ``;
	// 	assert.strictEqual(
	// 		jsonedData,
	// 		expectedHtml,
	// 		"Copied HTML should match for Footnote formula"
	// 	);
	// 	done();
	// });

	// QUnit.test("Paste mso styled text from word", function (assert) {
	// 	initDocument(logicDocument);
	// 	let done = assert.async();
	//
	// 	// htmlElement.setAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
	// 	// htmlElement.setAttribute("xmlns:w", "urn:schemas-microsoft-com:office:word");
	// 	// htmlElement.setAttribute('xmlns:m', "http://schemas.microsoft.com/office/2004/12/omml");
	// 	// htmlElement.setAttribute("xmlns", "http://www.w3.org/TR/REC-html40");
	//
	// 	const htmlElement = document.createElement("div");
	// 	htmlElement.innerHTML = `<html xmlns:v="urn:schemas-microsoft-com:vml"
	// 	xmlns:o="urn:schemas-microsoft-com:office:office"
	// 	xmlns:w="urn:schemas-microsoft-com:office:word"
	// 	xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
	// 	xmlns="http://www.w3.org/TR/REC-html40">
	//
	// 	<head>
	// 	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	// 	<meta name=ProgId content=Word.Document>
	// 	<meta name=Generator content="Microsoft Word 15">
	// 	<meta name=Originator content="Microsoft Word 15">
	// 	<link rel=File-List
	// 	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_filelist.xml">
	// 	<link rel=Edit-Time-Data
	// 	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_editdata.mso">
	// 	<link rel=themeData
	// 	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_themedata.thmx">
	// 	<link rel=colorSchemeMapping
	// 	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_colorschememapping.xml">
	// 	<style>
	// 	<!--
	// 	 /* Font Definitions */
	// 	 @font-face
	// 	\t{font-family:"Cambria Math";
	// 	\tpanose-1:2 4 5 3 5 4 6 3 2 4;
	// 	\tmso-font-charset:204;
	// 	\tmso-generic-font-family:roman;
	// 	\tmso-font-pitch:variable;
	// 	\tmso-font-signature:-536869121 1107305727 33554432 0 415 0;}
	// 	@font-face
	// 	\t{font-family:Aptos;
	// 	\tmso-font-charset:0;
	// 	\tmso-generic-font-family:swiss;
	// 	\tmso-font-pitch:variable;
	// 	\tmso-font-signature:536871559 3 0 0 415 0;}
	// 	 /* Style Definitions */
	// 	 p.MsoNormal, li.MsoNormal, div.MsoNormal
	// 	\t{mso-style-unhide:no;
	// 	\tmso-style-qformat:yes;
	// 	\tmso-style-parent:"";
	// 	\tmargin-top:0cm;
	// 	\tmargin-right:0cm;
	// 	\tmargin-bottom:8.0pt;
	// 	\tmargin-left:0cm;
	// 	\tline-height:115%;
	// 	\tmso-pagination:widow-orphan;
	// 	\tfont-size:12.0pt;
	// 	\tfont-family:"Aptos",sans-serif;
	// 	\tmso-ascii-font-family:Aptos;
	// 	\tmso-ascii-theme-font:minor-latin;
	// 	\tmso-fareast-font-family:Aptos;
	// 	\tmso-fareast-theme-font:minor-latin;
	// 	\tmso-hansi-font-family:Aptos;
	// 	\tmso-hansi-theme-font:minor-latin;
	// 	\tmso-bidi-font-family:"Times New Roman";
	// 	\tmso-bidi-theme-font:minor-bidi;
	// 	\tmso-font-kerning:1.0pt;
	// 	\tmso-ligatures:standardcontextual;
	// 	\tmso-fareast-language:EN-US;}
	// 	.MsoChpDefault
	// 	\t{mso-style-type:export-only;
	// 	\tmso-default-props:yes;
	// 	\tfont-family:"Aptos",sans-serif;
	// 	\tmso-ascii-font-family:Aptos;
	// 	\tmso-ascii-theme-font:minor-latin;
	// 	\tmso-fareast-font-family:Aptos;
	// 	\tmso-fareast-theme-font:minor-latin;
	// 	\tmso-hansi-font-family:Aptos;
	// 	\tmso-hansi-theme-font:minor-latin;
	// 	\tmso-bidi-font-family:"Times New Roman";
	// 	\tmso-bidi-theme-font:minor-bidi;
	// 	\tmso-fareast-language:EN-US;}
	// 	.MsoPapDefault
	// 	\t{mso-style-type:export-only;
	// 	\tmargin-bottom:8.0pt;
	// 	\tline-height:115%;}
	// 	@page WordSection1
	// 	\t{size:595.3pt 841.9pt;
	// 	\tmargin:2.0cm 42.5pt 2.0cm 3.0cm;
	// 	\tmso-header-margin:35.4pt;
	// 	\tmso-footer-margin:35.4pt;
	// 	\tmso-paper-source:0;}
	// 	div.WordSection1
	// 	\t{page:WordSection1;}
	// 	-->
	// 	</style>
	// 	</head>
	//
	// 	<body lang=RU style='tab-interval:35.4pt;word-wrap:break-word'>
	// 	<!--StartFragment-->
	//
	// 	<p class=MsoNormal><v:rect id="Р СѓРєРѕРїРёСЃРЅС‹Р№_x0020_РІРІРѕРґ_x0020_2" o:spid="_x0000_s1026"
	// 	 style='position:absolute;margin-left:230.8pt;margin-top:91.35pt;width:1.15pt;
	// 	 height:1.05pt;z-index:251660288;visibility:visible;mso-wrap-style:square;
	// 	 mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;
	// 	 mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;
	// 	 mso-position-horizontal-relative:text;mso-position-vertical:absolute;
	// 	 mso-position-vertical-relative:text' coordorigin=",1" coordsize="6,1"
	// 	 filled="f" strokecolor="#e71224" strokeweight=".35mm">
	// 	 <v:stroke endcap="round"/>
	// 	 <v:path shadowok="f" o:extrusionok="f" fillok="f" insetpenok="f"/>
	// 	 <o:lock v:ext="edit" rotation="t" aspectratio="t" verticies="t" text="t"
	// 	  shapetype="t"/>
	// 	 <o:ink i="AE0dAgYEARBYz1SK5pfFT48G+LrS4ZsiAwtIEETnpZABRSNGIwUDOAtkGQs4CQD+/wMAAAAAAAoW
	// 	AgRQAlAQX/9f/woAESBQvSN0PyncAS==
	// 	" annotation="t"/>
	// 	</v:rect><i style='mso-bidi-font-style:normal'><span lang=EN-US
	// 	style='mso-ansi-language:EN-US'>a<span style='color:#EE0000'>sdfasdfas</span><span
	// 	style='background:yellow;mso-highlight:yellow'>df</span><o:p></o:p></span></i></p>
	//
	// 	<!--EndFragment-->
	// 	</body>
	//
	// 	</html>`;
	//
	// 	AscTest.Editor.asc_PasteData(
	// 		AscCommon.c_oAscClipboardDataFormat.HtmlElement,
	// 		htmlElement
	// 	);
	//
	// 	logicDocument.SelectAll();
	// 	let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
	// 	oCopyProcessor.Start();
	// 	const copiedHtml = oCopyProcessor.getInnerHtml();
	// 	logicDocument.RemoveSelection();
	//
	// 	const jsonedData = removeBase64(JSON.stringify(copiedHtml));
	// 	const expectedHtml =
	// 		'"<p style=\\"line-height:13.8pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\" class=\\"docData;\\"><span style=\\"font-family:\'Times New Roman\';font-size:12pt;color:#000000;mso-style-textfill-fill-color:#000000\\"><i>a</i></span><span style=\\"font-family:\'Times New Roman\';font-size:12pt;color:#ee0000;mso-style-textfill-fill-color:#ee0000\\"><i>sdfasdfas</i></span><span style=\\"font-family:\'Times New Roman\';font-size:12pt;background-color:#ffff00;color:#000000;mso-style-textfill-fill-color:#000000\\"><i>df</i></span></p><p style=\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\">&nbsp;</p>"';
	// 	assert.strictEqual(
	// 		jsonedData,
	// 		expectedHtml,
	// 		"Copied HTML should match for mso styled text from word"
	// 	);
	// 	done();
	// });

	QUnit.module("Word Copy/Paste Tests");
});

function removeBase64(html) {
	// 1. Remove long base64-like strings (letters, digits, +, /, =)
	html = html.replace(/([A-Za-z0-9+/=]{50,})/g, "");

	// 2. Remove dynamic docData metadata like: docData;DOCY;v5;3707;
	html = html.replace(/docData;DOCY;v\d+;\d+;?/g, "docData;");

	return html;
}

function ToJsonString(logicDocument) {
	var oWriter = new AscJsonConverter.WriterToJSON();

	var oResult = {
		type: "document",
		textPr: logicDocument.GetText(),
		content: oWriter.SerContent(
			logicDocument.Content,
			undefined,
			undefined,
			undefined,
			true
		),
		// "paraPr":    bWriteDefaultParaPr ? oWriter.SerParaPr(this.GetDefaultParaPr().ParaPr) : undefined,
		// "theme":     bWriteTheme ? oWriter.SerTheme(this.Document.GetTheme()) : undefined,
		// "sectPr":    bWriteSectionPr ? oWriter.SerSectionPr(this.Document.SectPr) : undefined,
		// "numbering": bWriteNumberings ? oWriter.jsonWordNumberings : undefined,
		// "styles":    bWriteStyles ? oWriter.SerWordStylesForWrite() : undefined
	};

	return JSON.stringify(oResult);
}
