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
	AscTest.Editor.GetDocument = AscCommon.DocumentEditorApi.prototype.GetDocument.bind(AscTest.Editor);
	AscTest.Editor.ReplaceTextSmart = AscCommon.DocumentEditorApi.prototype.ReplaceTextSmart.bind(AscTest.Editor);

	function GetParagraphRunsInfo(paragraph)
	{
		let result = [];
		paragraph.CheckRunContent(function(run)
		{
			let text = run.GetText({
				ParaSeparator : "",
				Text : ""
			});
			if (!text || !text.length)
				return;
			
			let reviewType = run.GetReviewType();
			if (result.length && reviewType === result[result.length - 1][0])
				result[result.length - 1][1] += text;
			else
				result.push([text, lazyCopy(run.GetTextPr()), reviewType]);
		});
		return result;
	}

	function lazyCopy(obj) {
		const newObj = {};
		for (const [key, value] of Object.entries(obj)) {
			if ( typeof value !== 'function' && // no methods
				typeof value !== 'object' && // no objects
				value !== undefined) {
				newObj[key] = value;
			}
		}
		return newObj;
	}

	function setCurUser(sUserId) {
		let oDocInfo = new Asc.asc_CDocInfo();
		let oUserInfo = new Asc.asc_CUserInfo();

		oDocInfo.put_UserInfo(oUserInfo);
		oUserInfo.asc_putId(sUserId);

		AscTest.Editor.DocInfo = oDocInfo;
	}

	let logicDocument = AscTest.CreateLogicDocument();
	logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);
	
	QUnit.module("Test the Api.ReplaceTextSmart method");
	
	function fillDocument(runData)
	{
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		
		runData.forEach(function(d)
		{
			let run = AscTest.CreateRun();
			run.AddText(d.text);
			if (undefined !== d.reviewType)
				run.SetReviewType(d.reviewType);
			else
				run.SetReviewType(reviewtype_Common);
			
			if (d.pr) {
				let oTextPr = new AscWord.CTextPr();
				oTextPr.Set_FromObject(d.pr);
				run.Set_Pr(oTextPr);
			}

			p.AddToContentToEnd(run);
		});
		
		return p;
	}
	
	QUnit.test("Replace text without revisions", function (assert)
	{
		///////////////// 1 //////////////
		let p = fillDocument([
			{text: "E", pr: {Italic: true}},
			{text: "x"},
			{text: "a", pr: {Bold: true}},
			{text: "m"},
			{text: "p", pr: {Italic: true}},
			{text: "l", pr: {Bold: true}},
			{text: "e"},
			{text: " Test"}
		]);
		
		AscTest.SelectParagraph(p);
		AscTest.Editor.ReplaceTextSmart(["Sample Test"]);
		
		assert.deepEqual(
			GetParagraphRunsInfo(p),
			[
				["S", {Italic: true}, reviewtype_Common],
				["a", {Bold: true}, reviewtype_Common],
				["m", {}, reviewtype_Common],
				["p", {Italic: true}, reviewtype_Common],
				["l", {Bold: true}, reviewtype_Common],
				["e", {}, reviewtype_Common],
				[" Test", {}, reviewtype_Common],
			],
			"ReplaceTextSmart - complex multi-run replacement"
		);

		///////////////// 2 //////////////
		p = fillDocument([
			{text: "Start "},
			{text: "Bold", pr: {Bold: true}},
			{text: " and "},
			{text: "Italic", pr: {Italic: true}},
			{text: " End"},
		]);
		
		AscTest.SelectParagraph(p);
		AscTest.Editor.ReplaceTextSmart(["Start Strong and Soft End"]);
		
		assert.deepEqual(
			GetParagraphRunsInfo(p),
			[
				["Start ", {}, reviewtype_Common],
				["Strong", {Bold: true}, reviewtype_Common],
				[" and ", {}, reviewtype_Common],
				["Soft", {Italic: true}, reviewtype_Common],
				[" End", {}, reviewtype_Common],
			],
			"ReplaceTextSmart - spans across formatting boundaries"
		);

		///////////////// 3 //////////////
		p = fillDocument([
			{text: "User: "},
			{text: "admin123", pr: {Italic: true}},
			{text: " logged in"},
		]);
		
		AscTest.SelectParagraph(p);
		AscTest.Editor.ReplaceTextSmart(["User: guest logged in"]);
		
		assert.deepEqual(
			GetParagraphRunsInfo(p),
			[
				["User: ", {}, reviewtype_Common],
				["guest", {Italic: true}, reviewtype_Common],
				[" logged in", {}, reviewtype_Common],
			],
			"ReplaceTextSmart - replace middle text, preserve formatting"
		);

		///////////////// 4 //////////////
		p = fillDocument([
			{text: "Your code is "},
			{text: "WRONG", pr: {Bold: true}},
			{text: "!"},
		]);
		
		AscTest.SelectParagraph(p);
		AscTest.Editor.ReplaceTextSmart(["Your code is RIGHT!"]);
		
		assert.deepEqual(
			GetParagraphRunsInfo(p),
			[
				["Your code is ", {}, reviewtype_Common],
				["RIGHT", {Bold: true}, reviewtype_Common],
				["!", {}, reviewtype_Common],
			],
			"ReplaceTextSmart - replace within formatted and unformatted boundary"
		);
	});

	QUnit.test("Replace text with revisions", function (assert)
	{
		///////////////// 1 //////////////
		let p = fillDocument([
			{text: "full text added in review ", pr: {Italic: true}, reviewType: reviewtype_Add},
			{text: "removed in review ", pr: {Bold: true}, reviewType: reviewtype_Remove},
			{text: "common text", pr: {Italic: true}},
		]);
		
		AscTest.SelectParagraph(p);
		AscTest.Editor.ReplaceTextSmart(["text added in review and another interted text"]);
		
		assert.deepEqual(
			GetParagraphRunsInfo(p),
			[
				["text added in review ", {Italic: true}, reviewtype_Add],
				["and another interte", {Bold: true}, reviewtype_Remove],
				["d text", {Italic: true}, reviewtype_Common]
			],
			"ReplaceTextSmart - replace reviewed changes without track revisions"
		);

		///////////////// 2 //////////////
		AscTest.SetTrackRevisions(true);

		p = fillDocument([
			{text: "full text added in review ", pr: {Italic: true}, reviewType: reviewtype_Add},
			{text: "removed in review ", pr: {Bold: true}, reviewType: reviewtype_Remove},
			{text: "common text", pr: {Italic: true}},
		]);
		
		AscTest.SelectParagraph(p);
		AscTest.Editor.ReplaceTextSmart(["text added in review and another interted text"]);
		
		assert.deepEqual(
			GetParagraphRunsInfo(p),
			[
				["text added in review ", {Italic: true}, reviewtype_Add],
				["removed in review ", {Bold: true}, reviewtype_Remove],
				["c", {Italic: true}, reviewtype_Remove],
				["and an", {Italic: true}, reviewtype_Add],
				["o", {Italic: true}, reviewtype_Common],
				["mmo", {Italic: true}, reviewtype_Remove],
				["ther i", {Italic: true}, reviewtype_Add],
				["nterted text", {Italic: true}, reviewtype_Common]
			],
			"ReplaceTextSmart - replace reviewed changes with track revisions"
		);

		///////////////// 3 //////////////
		p = fillDocument([
			{text: "text removed by first user and ", pr: {Bold: true}, reviewType: reviewtype_Remove},
			{text: "common text and ", pr: {Italic: true}},
			{text: "text added by first user", pr: {Bold: true}, reviewType: reviewtype_Add},
		]);

		// change user
		setCurUser('uid-2');

		AscTest.SelectParagraph(p);
		AscTest.Editor.ReplaceTextSmart(["edited common text"]);
		
		assert.deepEqual(
			GetParagraphRunsInfo(p),
			[
				["text removed by first user and ", {Bold: true}, reviewtype_Remove],
				["edited commo", {Italic: true}, reviewtype_Common],
				["n text a", {Italic: true}, reviewtype_Remove],
				["n", {Italic: true}, reviewtype_Common],
				["d", {Italic: true}, reviewtype_Remove],
				[" ", {Italic: true}, reviewtype_Common],
				["tex", {Bold: true}, reviewtype_Add],
				["t added by firs", {Bold: true}, reviewtype_Remove],
				["t", {Bold: true}, reviewtype_Add],
				[" user", {Bold: true}, reviewtype_Remove]
			],
			"ReplaceTextSmart - replace reviewed changes with another user and track revisions"
		);
	});
});
