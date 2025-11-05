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

	QUnit.test("Test: \"callback tests paste plain text\"", function (assert) {
		let done = assert.async();
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, "test", undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste HTML\"", function (assert) {
		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "test HTML content";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste Internal format\"", function (assert) {
		let done = assert.async();
		let binaryData = "";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, binaryData, undefined, undefined, undefined, function () {
			assert.ok(true);
			done();
		});
	});

	QUnit.test('Test: "copy sdt email address HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		// Create an HTML element to simulate copying
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<span>\n" +
			"    <w:Sdt Title=\"Email Address\" Form=\"t\" Key=\"EmailAddress1\" Border=\"red\" Shd=\"blue\" HelpText=\"Enter your email\" Required=\"t\" RoleName=\"Recipient\" RoleColor=\"#7FB5B5\" sdttag=\"EmailTag\" Label=\"Email\" ID=\"1837335015\">\n" +
			"        <w:TextForm MaxCharacters=\"-1\" Comb=\"f\" WidthRule=\"1\" MultiLine=\"f\" AutoFit=\"f\" FormatType=\"regexp\" RegExp=\"\\S+@\\S+\\.\\S+\" Symbols=\"49,50,51\"/>\n" +
			"        user111@example.com\n" +
			"    </w:Sdt>\n" +
			"</span>\n";

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
			"type": "document",
			"textPr": " user111@example.com \r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "run"
						},
						{
							"sdtPr": {
								"alias": "Email Address",
								"appearance": "boundingBox",
								"equation": false,
								"id": "1837335015",
								"picture": false,
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": false,
								"tag": "EmailTag",
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										" user111@example.com "
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt checkbox form HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<p style=''>\n" +
			"    <w:Sdt ShowingPlcHdr=\"t\" CheckBox=\"t\" CheckBoxIsChecked=\"f\" CheckBoxValueChecked=\"☑\" CheckBoxValueUnchecked=\"☐\" CheckBoxFontChecked=\"Segoe UI Symbol\" CheckBoxFontUnchecked=\"Segoe UI Symbol\" Title=\"Title\" Form=\"t\" Key=\"DropDown1\" Border=\"red\" Shd=\"blue\" HelpText=\"HelpText\" Required=\"t\" RoleName=\"RoleName\" RoleColor=\"#7FB5B5\" sdttag=\"Tag\" Label=\"Label\" ID=\"-1395741881\">\n" +
			"        <span style='font-family:\"Segoe UI Symbol\",sans-serif'>☐</span>\n" +
			"    </w:Sdt>\n" +
			"</p>\n";

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
			"type": "document",
			"textPr": "☐ \r\n\r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"pBdr": {
							"bottom": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"left": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"right": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"top": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							}
						},
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"sdtPr": {
								"alias": "Title",
								"appearance": "boundingBox",
								"equation": false,
								"id": "-1395741881",
								"picture": false,
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": false,
								"tag": "Tag",
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										"☐ "
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				},
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt combobox form HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<span>\n" +
			"    <w:Sdt Title=\"Title\" Form=\"t\" Key=\"ComboBox2\" Border=\"red\" Shd=\"blue\" HelpText=\"HelpText\" Required=\"t\" RoleName=\"RoleName\" RoleColor=\"#7FB5B5\" sdttag=\"Tag\" Label=\"Label\" ComboBox=\"t\" ID=\"1837335014\">\n" +
			"        <w:ListItem ListValue=\"Choose an item1\" DataValue=\"Choose an item1\"/>test\n" +
			"    </w:Sdt>\n" +
			"</span>\n";

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
			"type": "document",
			"textPr": "test \r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "run"
						},
						{
							"sdtPr": {
								"alias": "Title",
								"appearance": "boundingBox",
								"comboBox": {
									"listItem": [],
									"lastValue": -1
								},
								"equation": false,
								"id": "1837335014",
								"picture": false,
								"placeholder": {
									"docPart": "DefaultPlaceholder_LIST"
								},
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": true,
								"tag": "Tag",
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										"test "
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt text inline form HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<span>\n" +
			"    <w:Sdt Title=\"Title\" Form=\"t\" Key=\"DropDown1\" Border=\"red\" Shd=\"blue\" HelpText=\"HelpText\" Required=\"t\" RoleName=\"RoleName\" RoleColor=\"#7FB5B5\" sdttag=\"Tag\" Label=\"Label\" ID=\"1837335014\">\n" +
			"        inline\n" +
			"    </w:Sdt>\n" +
			"</span>\n";

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
			"type": "document",
			"textPr": " inline \r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "run"
						},
						{
							"sdtPr": {
								"alias": "Title",
								"appearance": "boundingBox",
								"equation": false,
								"id": "1837335014",
								"picture": false,
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": false,
								"tag": "Tag",
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										" inline "
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt radio button form HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<p style=''>\n" +
			"    <w:Sdt CheckBox=\"t\" CheckBoxIsChecked=\"f\" CheckBoxValueChecked=\"◙\" CheckBoxValueUnchecked=\"○\" CheckBoxFontChecked=\"Segoe UI Symbol\" CheckBoxFontUnchecked=\"Segoe UI Symbol\" GroupKey=\"Group 1\" Form=\"t\" Key=\"Choice1\" ContentLocked=\"t\" ID=\"123456789\">Choice 1</w:Sdt>\n" +
			"</p>\n";

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
			"type": "document",
			"textPr": "Choice 1\r\n\r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"pBdr": {
							"bottom": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"left": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"right": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"top": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							}
						},
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"sdtPr": {
								"appearance": "boundingBox",
								"equation": false,
								"id": "123456789",
								"lock": "sdtContentLocked",
								"picture": false,
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": false,
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										"Choice 1"
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				},
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt date picker form HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<span>\n" +
			"    <w:Sdt DocPart=\"0D4FD865761947FCBA0D9229E17016DB\" Calendar=\"t\" MapToDateTime=\"t\" CalendarType=\"Gregorian\" Date=\"2022-10-24\" DateFormat=\"dd.MM.yyyy\" Lang=\"EN-US\" Title=\"Date Picker\" Form=\"t\" Key=\"DatePicker1\" Border=\"blue\" Shd=\"yellow\" HelpText=\"Select a date\" Required=\"t\" RoleName=\"DateRole\" RoleColor=\"#FF5733\" sdttag=\"DateTag\" Label=\"DateLabel\" ID=\"-291673853\">24.10.2022</w:Sdt>\n" +
			"</span>\n";

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
			"type": "document",
			"textPr": "24.10.2022\r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "run"
						},
						{
							"sdtPr": {
								"alias": "Date Picker",
								"appearance": "boundingBox",
								"date": {
									"calendar": "gregorian",
									"dateFormat": "dd.MM.yyyy",
									"lid": "en-US",
									"fullDate": "2022-10-24T00:00:00Z"
								},
								"equation": false,
								"id": "-291673853",
								"picture": false,
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": false,
								"tag": "DateTag",
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										"24.10.2022"
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt complex form HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<p style=''>\n" +
			"    <w:Sdt ComplexForm=\"t\" ComplexFormType=\"0\" Title=\"\" sdttag=\"tag\" Label=\"\" ContentLocked=\"t\" ShowingPlcHdr=\"t\" Temporary=\"f\" Form=\"t\" Key=\"Complex1\" HelpText=\"tip\" Fixed=\"f\" ID=\"123456789\">\n" +
			"        Complex form content\n" +
			"    </w:Sdt>\n" +
			"</p>\n";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
			undefined,
			undefined,
			undefined,
			function () {}
		);

		const result = ToJsonString(logicDocument).replace(/"docPart":"[^"]*"/g, '"docPart":""');

		const expected = {
			"type": "document",
			"textPr": " Complex form content \r\n\r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"pBdr": {
							"bottom": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"left": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"right": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"top": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							}
						},
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"sdtPr": {
								"appearance": "boundingBox",
								"equation": false,
								"id": "123456789",
								"lock": "sdtContentLocked",
								"picture": false,
								"placeholder": {
									"docPart": ""
								},
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": true,
								"tag": "tag",
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										" Complex form content "
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				},
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt text form with mask HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<span>\n" +
			"    <w:Sdt Form=\"t\" Key=\"CreditCard1\" HelpText=\"Enter credit card number\" Required=\"t\" sdttag=\"CreditCardTag\" Label=\"CreditCardLabel\" ID=\"1837335025\" PlcHdr=\"PlaceholderText\" showingplchdr=\"t\">\n" +
			"        <w:TextForm MaxCharacters=\"-1\" Comb=\"f\" WidthRule=\"1\" MultiLine=\"f\" AutoFit=\"f\" FormatType=\"mask\" Mask=\"9999-9999-9999-9999\"/>\n" +
			"        1234-5678-9012-3456\n" +
			"    </w:Sdt>\n" +
			"</span>\n";

		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement,
			undefined,
			undefined,
			undefined,
			function () {}
		);

		const result = ToJsonString(logicDocument).replace(/"docPart":"[^"]*"/g, '"docPart":""');;

		const expected = {
			"type": "document",
			"textPr": " 1234-5678-9012-3456 \r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "run"
						},
						{
							"sdtPr": {
								"appearance": "boundingBox",
								"equation": false,
								"id": "1837335025",
								"picture": false,
								"placeholder": {
									"docPart": ""
								},
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": true,
								"tag": "CreditCardTag",
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										" 1234-5678-9012-3456 "
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt plain text HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<span>\n" +
			"    <w:Sdt DocPart=\"AD334DF7FD99465099BBBD9F3C3AE5AE\" ID=\"-818574957\">\n" +
			"        <span style='mso-spacerun:yes'> </span>\n" +
			"        Plain text\n" +
			"    </w:Sdt>\n" +
			"</span>\n";

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
			"type": "document",
			"textPr": "  Plain text \r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "run"
						},
						{
							"sdtPr": {
								"appearance": "boundingBox",
								"equation": false,
								"id": "-818574957",
								"picture": false,
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"showingPlcHdr": false,
								"temporary": false
							},
							"content": [
								{
									"bFromDocument": true,
									"rPr": {
										"bFromDocument": true,
										"type": "textPr"
									},
									"content": [
										"  Plain text "
									],
									"footnotes": [],
									"endnotes": [],
									"reviewType": "common",
									"type": "run"
								}
							],
							"type": "inlineLvlSdt"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

	QUnit.test('Test: "copy sdt rich text block level HTML with JSON verification"', function (assert) {
		initDocument();

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML =
			"<w:Sdt DocPart=\"5AAA70241BB24C68A16AC32AB3233FFC\" ID=\"819472493\">\n" +
			"    <p class=\"MsoNormal\">Rich plain text\n" +
			"        <o:p/>\n" +
			"        <w:sdtPr/>\n" +
			"    </p>\n" +
			"</w:Sdt>\n";

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
			"type": "document",
			"textPr": "\r\nRich plain text \r\n\r\n",
			"content": [
				{
					"bFromDocument": true,
					"pPr": {
						"pBdr": {
							"bottom": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"left": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"right": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							},
							"top": {
								"color": {
									"auto": false,
									"r": 0,
									"g": 0,
									"b": 0
								},
								"sz": 4,
								"space": 0,
								"value": "none"
							}
						},
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "run"
						},
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				},
				{
					"sdtPr": {
						"appearance": "boundingBox",
						"equation": false,
						"id": "819472493",
						"picture": false,
						"rPr": {
							"bFromDocument": true,
							"type": "textPr"
						},
						"showingPlcHdr": false,
						"temporary": false
					},
					"sdtContent": {
						"bPresentation": false,
						"content": [
							{
								"bFromDocument": true,
								"pPr": {
									"pBdr": {
										"bottom": {
											"color": {
												"auto": false,
												"r": 0,
												"g": 0,
												"b": 0
											},
											"sz": 4,
											"space": 0,
											"value": "none"
										},
										"left": {
											"color": {
												"auto": false,
												"r": 0,
												"g": 0,
												"b": 0
											},
											"sz": 4,
											"space": 0,
											"value": "none"
										},
										"right": {
											"color": {
												"auto": false,
												"r": 0,
												"g": 0,
												"b": 0
											},
											"sz": 4,
											"space": 0,
											"value": "none"
										},
										"top": {
											"color": {
												"auto": false,
												"r": 0,
												"g": 0,
												"b": 0
											},
											"sz": 4,
											"space": 0,
											"value": "none"
										}
									},
									"bFromDocument": true,
									"type": "paraPr"
								},
								"rPr": {
									"bFromDocument": true,
									"type": "textPr"
								},
								"content": [
									{
										"bFromDocument": true,
										"rPr": {
											"bFromDocument": true,
											"type": "textPr"
										},
										"content": [
											"Rich plain text "
										],
										"footnotes": [],
										"endnotes": [],
										"reviewType": "common",
										"type": "run"
									},
									{
										"bFromDocument": true,
										"rPr": {
											"bFromDocument": true,
											"type": "textPr"
										},
										"content": [],
										"footnotes": [],
										"endnotes": [],
										"reviewType": "common",
										"type": "endRun"
									}
								],
								"changes": [],
								"type": "paragraph"
							}
						],
						"type": "docContent"
					},
					"type": "blockLvlSdt"
				},
				{
					"bFromDocument": true,
					"pPr": {
						"bFromDocument": true,
						"type": "paraPr"
					},
					"rPr": {
						"bFromDocument": true,
						"type": "textPr"
					},
					"content": [
						{
							"bFromDocument": true,
							"rPr": {
								"bFromDocument": true,
								"type": "textPr"
							},
							"content": [],
							"footnotes": [],
							"endnotes": [],
							"reviewType": "common",
							"type": "endRun"
						}
					],
					"changes": [],
					"type": "paragraph"
				}
			]
		};

		assert.strictEqual(
			result,
			JSON.stringify(expected),
			"HTML content should match expected JSON format"
		);

		done();
	});

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


	QUnit.module("Word Copy Paste Tests");
});
