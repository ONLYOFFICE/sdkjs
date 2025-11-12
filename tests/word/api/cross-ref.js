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
	
	QUnit.module("Test adding and working with cross-references");
	
	QUnit.test("Test adding cross-ref to a block-level sdt", function (assert)
	{
		AscTest.ClearDocument();
		
		let cc = AscTest.CreateBlockLvlSdt();
		logicDocument.PushToContent(cc);
		
		let docContent = cc.GetContent();
		docContent.ClearContent(false);
		let headingParagraph = AscTest.CreateParagraph();
		docContent.AddToContent(0, headingParagraph);
		
		let styleId = styleManager.GetDefaultHeading(0);
		headingParagraph.SetParagraphStyle(styleManager.Get(styleId).GetName());
		AscTest.AddTextToParagraph(headingParagraph, "HeadingText");
		
		assert.strictEqual(AscTest.GetParagraphText(headingParagraph), "HeadingText", "Check paragraph text");
		
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		AscTest.MoveCursorToParagraph(p);
		
		logicDocument.AddRefToParagraph(headingParagraph, 0, true, false, undefined);
		assert.strictEqual(AscTest.GetParagraphText(p), "HeadingText", "Check text after adding ref to a block-level sdt");
		assert.strictEqual(headingParagraph.GetBookmarkRefToParagraph(), "_Ref1", "Check bookmark name");
		
		// Check bug 69293. Adding a cross-ref to a locked block-level sdt
		
		AscTest.ClearDocument();
		
		cc = AscTest.CreateBlockLvlSdt();
		logicDocument.PushToContent(cc);
		
		docContent = cc.GetContent();
		docContent.ClearContent(false);
		headingParagraph = AscTest.CreateParagraph();
		docContent.AddToContent(0, headingParagraph);
		
		headingParagraph.SetParagraphStyle(styleManager.Get(styleId).GetName());
		AscTest.AddTextToParagraph(headingParagraph, "HeadingText");
		
		cc.SetContentControlLock(Asc.c_oAscSdtLockType.SdtContentLocked);
		
		p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		AscTest.MoveCursorToParagraph(p);
		
		logicDocument.AddRefToParagraph(headingParagraph, 0, true, false, undefined);
		assert.strictEqual(AscTest.GetParagraphText(p), "HeadingText", "Check text after adding ref to a locked block-level sdt");
		assert.strictEqual(headingParagraph.GetBookmarkRefToParagraph(), "_Ref1", "Check bookmark name");
	});
});
