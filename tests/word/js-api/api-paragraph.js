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
	QUnit.module('Test the ApiParagraph methods');

	function createApiParagraph() {
		return AscTest.JsApi.CreateParagraph();
	}

	QUnit.test('SetShd, GetShd', function (assert) {
		const apiParagraph = createApiParagraph();

		assert.strictEqual(apiParagraph.GetShd(), null, 'Shading check for a newly created paragraph');

		apiParagraph.SetShd('clear', 255, 122, 100);
		assert.equalRgb(apiParagraph.GetShd(), { r: 255, g: 122, b: 100 }, 'Check shd color set with RGB components');

		apiParagraph.SetShd('clear', AscTest.JsApi.HexColor('55aa00'));
		assert.equalRgb(apiParagraph.GetShd(), { r: 85, g: 170, b: 0 }, 'Check shd color set with ApiColor (hex)');

		apiParagraph.SetShd('clear', AscTest.JsApi.ThemeColor('accent2'));
		assert.strictEqual(apiParagraph.GetShd().IsThemeColor(), true, 'Check shd color set with ApiColor (theme)');

		apiParagraph.SetShd('clear', AscTest.JsApi.AutoColor());
		assert.strictEqual(apiParagraph.GetShd().IsAutoColor(), true, 'Check shd color set with ApiColor (auto)');
	});

	QUnit.test('SetColor, GetColor', function (assert) {
		const hexColor = AscTest.JsApi.HexColor('#bada55');
		const themeColor = AscTest.JsApi.ThemeColor('accent2');
		const autoColor = AscTest.JsApi.AutoColor();

		const apiParagraph = createApiParagraph();
		apiParagraph.AddText('Run for testing paragraph color');

		let apiRun;

		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor(), null, 'Color check for a newly created paragraph');

		apiParagraph.SetColor(80, 160, 240);
		apiRun = apiParagraph.GetElement(0);
		assert.equalRgb(apiRun.GetColor(), { r: 80, g: 160, b: 240 }, 'Color check after setting color with RGB components');

		apiParagraph.SetColor(hexColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().GetHex(), '#bada55', 'Color check after setting color with ApiColor (hex)');

		apiParagraph.SetColor(themeColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().IsThemeColor(), true, 'Color check after setting color with ApiColor (theme)');

		apiParagraph.SetColor(autoColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().IsAutoColor(), true, 'Color check after setting color with ApiColor (auto)');
	});
});
