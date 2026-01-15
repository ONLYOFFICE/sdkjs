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
	QUnit.module('Test the ApiRange methods');
	
	function getFirstDocParagraph() {
		const doc = AscTest.JsApi.GetDocument();
		let par = doc.GetElement(0);
		if (par)
			return par;

		par = AscTest.JsApi.CreateParagraph();
		doc.Push(par);
		return par;
	}
	
	QUnit.test('GetClassType', function (assert)
	{
		const para = getFirstDocParagraph();
		para.AddText('Test text for range');
		const range = para.GetRange();
		
		assert.strictEqual(range.GetClassType(), 'range', 'Check range class type');
	});
	
	QUnit.test('GetText/AddText', function (assert)
	{
		let p = getFirstDocParagraph();
		let runBefore = p.AddText("QQQ");
		let run = p.AddText("1");
		run.AddTabStop();
		run.AddText("2");
		run.AddLineBreak();
		run.AddText("3");
		
		let range = run.GetRange();
		
		assert.strictEqual(range.GetText(), "1\t2\r3", "Check text");
		assert.strictEqual(range.GetText({
			"TabSymbol" : "_t_",
			"NewLineSeparator" : "_nl_",
		}), "1_t_2_nl_3", "Check text");
		
		range.AddText("Before", "before");
		range.AddText("After", "after");
		
		assert.strictEqual(p.GetText(), "QQQBefore1\t2\r3After\r\n", "Check paragraph text");
		assert.strictEqual(range.GetText(), "Before1\t2\r3After", "Check range text");
		
		runBefore = AscTest.JsApi.CreateRun();
		p.AddElement(runBefore, 0);
		runBefore.AddText("123");
		
		assert.strictEqual(p.GetText(), "123QQQBefore1\t2\r3After\r\n", "Check paragraph text");
		assert.strictEqual(range.GetText(), "Before1\t2\r3After", "Check range text");
	});
	
	QUnit.test('GetParagraph', function (assert)
	{
		const para = getFirstDocParagraph();
		para.AddText('Test paragraph text');
		const range = para.GetRange();
		
		const firstPara = range.GetParagraph(0);
		assert.strictEqual(firstPara.private_GetImpl(), para.private_GetImpl(), 'Check para range paragraph');
		
		const invalidPara = range.GetParagraph(10);
		assert.strictEqual(invalidPara, null, 'GetParagraph with invalid index should return null');
		
		const negativePara = range.GetParagraph(-1);
		assert.strictEqual(negativePara, null, 'GetParagraph with negative index should return null');
		
		let run = para.AddText("123");
		let runRange = run.GetRange();
		assert.strictEqual(runRange.GetRange().GetParagraph().private_GetImpl(), para.private_GetImpl(), 'Check run range paragraph');
	});
	
	QUnit.test('GetAllParagraphs', function (assert)
	{
		const doc = AscTest.JsApi.GetDocument();
		doc.RemoveAllElements();
		
		// Add first paragraph
		const para1 = AscTest.JsApi.CreateParagraph();
		para1.AddText('First paragraph');
		doc.Push(para1);
		
		// Add second paragraph
		const para2 = AscTest.JsApi.CreateParagraph();
		para2.AddText('Second paragraph');
		doc.Push(para2);
		
		// Add third paragraph
		const para3 = AscTest.JsApi.CreateParagraph();
		para3.AddText('Third paragraph');
		doc.Push(para3);
		
		// Test range for single paragraph
		const singleRange = para1.GetRange();
		const singleParagraphs = singleRange.GetAllParagraphs();
		assert.ok(Array.isArray(singleParagraphs), 'GetAllParagraphs should return an array');
		assert.strictEqual(singleParagraphs.length, 1, 'Single paragraph range should contain 1 paragraph');
		assert.strictEqual(singleParagraphs[0].private_GetImpl(), para1.private_GetImpl(), 'Check single paragraph range');
		
		// Test range for entire document
		const docRange = doc.GetRange();
		const allParagraphs = docRange.GetAllParagraphs();
		assert.ok(Array.isArray(allParagraphs), 'Document range should return an array');
		assert.strictEqual(allParagraphs.length, 3, 'Document range should contain 3 paragraphs');
		
		// Verify each paragraph
		assert.strictEqual(allParagraphs[0].private_GetImpl(), para1.private_GetImpl(), 'Check first paragraph in document range');
		assert.strictEqual(allParagraphs[1].private_GetImpl(), para2.private_GetImpl(), 'Check second paragraph in document range');
		assert.strictEqual(allParagraphs[2].private_GetImpl(), para3.private_GetImpl(), 'Check third paragraph in document range');
		
		// Test range spanning multiple paragraphs
		const para2Range = para2.GetRange();
		const para3Range = para3.GetRange();
		const multiRange = para2Range.ExpandTo(para3Range);
		const multiParagraphs = multiRange.GetAllParagraphs();
		assert.strictEqual(multiParagraphs.length, 2, 'Multi-paragraph range should contain 2 paragraphs');
		assert.strictEqual(multiParagraphs[0].private_GetImpl(), para2.private_GetImpl(), 'Check first paragraph in multi range');
		assert.strictEqual(multiParagraphs[1].private_GetImpl(), para3.private_GetImpl(), 'Check second paragraph in multi range');
	});

	QUnit.test('SetColor, GetColor', function (assert) {
		const rgbColor = AscTest.JsApi.RGB(255, 127, 0);
		const hexColor = AscTest.JsApi.HexColor('#bada55');
		const themeColor = AscTest.JsApi.ThemeColor('accent2');
		const autoColor = AscTest.JsApi.AutoColor();

		const apiParagraph = getFirstDocParagraph();
		apiParagraph.AddText('Paragraph for testing range color');
		const apiRange = apiParagraph.GetRange();

		let apiRun;

		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor(), null, 'Color check for an empty run');

		apiRange.SetColor(80, 160, 240);
		apiRun = apiParagraph.GetElement(0);
		assert.equalRgb(apiRun.GetColor(), { r: 80, g: 160, b: 240 }, 'Color check after setting color for range with RGB components');

		apiRange.SetColor(rgbColor);
		apiRun = apiParagraph.GetElement(0);
		assert.equalRgb(apiRun.GetColor(), { r: 255, g: 127, b: 0 }, 'Color check after setting color for range with ApiColor (RGB)');

		apiRange.SetColor(hexColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().GetHex(), '#BADA55', 'Color check after setting color for range with ApiColor (hex)');

		apiRange.SetColor(themeColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().IsThemeColor(), true, 'Color check after setting color for range with ApiColor (theme)');

		apiRange.SetColor(autoColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().IsAutoColor(), true, 'Color check after setting color for range with ApiColor (auto)');
	});

	QUnit.test('SetShd, GetShd', function (assert) {
		const rgbColor = AscTest.JsApi.RGB(255, 127, 0);
		const hexColor = AscTest.JsApi.HexColor('#bada55');
		const themeColor = AscTest.JsApi.ThemeColor('accent2');
		const autoColor = AscTest.JsApi.AutoColor();

		const apiParagraph = getFirstDocParagraph();
		apiParagraph.AddText('Paragraph for testing range color');
		const apiRange = apiParagraph.GetRange();

		let firstParagraph;

		firstParagraph = getFirstDocParagraph();
		assert.strictEqual(firstParagraph.GetShd(), null, 'Shading (Shd) check for a run with no shading');

		apiRange.SetShd('clear', 80, 160, 240);
		firstParagraph = getFirstDocParagraph();
		assert.equalRgb(firstParagraph.GetShd(), { r: 80, g: 160, b: 240 }, 'Shading check after setting shading for range with RGB components');

		apiRange.SetShd('clear', rgbColor);
		firstParagraph = getFirstDocParagraph();
		assert.equalRgb(firstParagraph.GetShd(), { r: 255, g: 127, b: 0 }, 'Shading check after setting shading for range with ApiColor (RGB)');

		apiRange.SetShd('clear', hexColor);
		firstParagraph = getFirstDocParagraph();
		assert.strictEqual(firstParagraph.GetShd().GetHex(), '#BADA55', 'Shading check after setting shading for range with ApiColor (hex)');

		apiRange.SetShd('clear', themeColor);
		firstParagraph = getFirstDocParagraph();
		assert.strictEqual(firstParagraph.GetShd().IsThemeColor(), true, 'Shading check after setting shading for range with ApiColor (theme)');

		apiRange.SetShd('clear', autoColor);
		firstParagraph = getFirstDocParagraph();
		assert.strictEqual(firstParagraph.GetShd().IsAutoColor(), true, 'Shading check after setting shading for range with ApiColor (auto)');
	});
});
