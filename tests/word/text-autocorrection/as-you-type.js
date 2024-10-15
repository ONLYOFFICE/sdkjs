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

$(function ()
{
	AscCommon.CFirstLetterExceptions.prototype.GetMaxLen = function()
	{
		return 7;
	}
	
	let logicDocument = AscTest.CreateLogicDocument();
	let p;
	
	QUnit.module("Test 'as you type' autocorrections");
	
	QUnit.testStart(function()
	{
		AscTest.ClearDocument();
	});
	
	function enterAndCheckText(assert, text, checkText)
	{
		p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.EnterText(text);
		AscTest.PressKey(AscTest.Key.space);
		assert.strictEqual(AscTest.GetParagraphText(p), checkText + " ", text + " -> " + checkText);
	}
	
	QUnit.test("Test: capitalize first letter of the sentence", function (assert)
	{
		logicDocument.SetAutoCorrectFirstLetterOfSentences(true);
		enterAndCheckText(assert, "hello", "Hello");
		enterAndCheckText(assert, "привет", "Привет");
		enterAndCheckText(assert, "გამარჯობა", "გამარჯობა"); // bug 69089
		enterAndCheckText(assert, "Hello world! hello", "Hello world! Hello");
		enterAndCheckText(assert, "Hello world! hello world", "Hello world! hello world");
	});
	
	QUnit.test("Test: capitalize first letter of the sentence in table and capitalize first letter of table cell", function (assert)
	{
		let table = AscTest.CreateTable(3, 3);
		logicDocument.AddToContent(0, table);
		let dc = table.GetRow(0).GetCell(0).GetContent();
		let p;
		
		function enterAndCheckTextToCell(assert, text, checkText)
		{
			dc.ClearContent(false);
			p = AscTest.CreateParagraph();
			dc.AddToContent(0, p);
			AscTest.MoveCursorToParagraph(p, true);
			AscTest.EnterText(text);
			AscTest.PressKey(AscTest.Key.space);
			assert.strictEqual(AscTest.GetParagraphText(p), checkText + " ", text + " -> " + checkText);
		}
		
		logicDocument.SetAutoCorrectFirstLetterOfCells(false);
		logicDocument.SetAutoCorrectFirstLetterOfSentences(false);
		enterAndCheckTextToCell(assert, "hello", "hello");
		enterAndCheckTextToCell(assert, "привет", "привет");
		enterAndCheckTextToCell(assert, "Hello world! hello", "Hello world! hello");
		enterAndCheckTextToCell(assert, "Hello world! hello world", "Hello world! hello world");
		
		logicDocument.SetAutoCorrectFirstLetterOfCells(true);
		logicDocument.SetAutoCorrectFirstLetterOfSentences(true);
		enterAndCheckTextToCell(assert, "hello", "Hello");
		enterAndCheckTextToCell(assert, "привет", "Привет");
		enterAndCheckTextToCell(assert, "Hello world! hello", "Hello world! Hello");
		enterAndCheckTextToCell(assert, "Hello world! hello world", "Hello world! hello world");
		
		logicDocument.SetAutoCorrectFirstLetterOfCells(false);
		logicDocument.SetAutoCorrectFirstLetterOfSentences(true);
		enterAndCheckTextToCell(assert, "hello", "hello");
		enterAndCheckTextToCell(assert, "привет", "привет");
		enterAndCheckTextToCell(assert, "Hello world! hello", "Hello world! Hello");
		enterAndCheckTextToCell(assert, "Hello world! hello world", "Hello world! hello world");
		
		logicDocument.SetAutoCorrectFirstLetterOfCells(true);
		logicDocument.SetAutoCorrectFirstLetterOfSentences(false);
		enterAndCheckTextToCell(assert, "hello", "Hello");
		enterAndCheckTextToCell(assert, "привет", "Привет");
		enterAndCheckTextToCell(assert, "Hello world! hello", "Hello world! hello");
		enterAndCheckTextToCell(assert, "Hello world! hello world", "Hello world! hello world");
	});
});
