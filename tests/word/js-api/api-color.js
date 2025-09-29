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
	QUnit.module('Test the ApiColor methods');
	
	QUnit.test('GetClassType, IsAutoColor, IsThemeColor', function(assert)
	{
		const autoColor  = AscTest.JsApi.AutoColor();
		const themeColor = AscTest.JsApi.ThemeColor('accent1');
		const hexColor   = AscTest.JsApi.HexColor('#bada55');
		
		assert.strictEqual(hexColor.GetClassType(), 'color', 'Class type of ApiColor instance');
		
		assert.strictEqual(hexColor.IsAutoColor(), false, 'HexColor should not be auto color');
		assert.strictEqual(autoColor.IsAutoColor(), true, 'AutoColor must be auto color');
		
		assert.strictEqual(hexColor.IsThemeColor(), false, 'HexColor should not have "theme" type');
		assert.strictEqual(themeColor.IsThemeColor(), true, 'ThemeColor must have "theme" type');
	});
	
	QUnit.test('GetRGB, GetRGBA, GetHex', function(assert)
	{
		const rgbColor   = AscTest.JsApi.RGB(186, 218, 85);
		const rgbaColor  = AscTest.JsApi.RGBA(186, 218, 85, 123);
		const hexColor   = AscTest.JsApi.HexColor('#bada55');
		const themeColor = AscTest.JsApi.ThemeColor('accent2');
		const autoColor  = AscTest.JsApi.AutoColor();
		
		assert.equalRgb(rgbColor.GetRGB(), {r : 186, g : 218, b : 85}, 'Check RGB from RGB color');
		assert.equalRgba(rgbColor.GetRGBA(), {r : 186, g : 218, b : 85, a : 255}, 'Check RGBA from RGBA color');
		assert.strictEqual(rgbColor.GetHex(), '#bada55', 'Check Hex from RGB color');
		
		assert.equalRgb(rgbaColor.GetRGB(), {r : 186, g : 218, b : 85}, 'Check RGB from RGBA color');
		assert.equalRgba(rgbaColor.GetRGBA(), {r : 186, g : 218, b : 85, a : 123}, 'Check RGBA from RGBA color');
		assert.strictEqual(rgbaColor.GetHex(), '#bada55', 'Check Hex from RGBA color');
		
		assert.equalRgb(hexColor.GetRGB(), {r : 186, g : 218, b : 85}, 'Check RGB from Hex color');
		assert.equalRgba(hexColor.GetRGBA(), {r : 186, g : 218, b : 85, a : 255}, 'Check RGBA from Hex color');
		assert.strictEqual(hexColor.GetHex(), '#bada55', 'Check Hex from Hex color');
		
		assert.equalRgb(themeColor.GetRGB(), {r : 192, g : 80, b : 77}, 'Check RGB from ThemeColor');
		assert.equalRgba(themeColor.GetRGBA(), {r : 192, g : 80, b : 77, a : 255}, 'Check RGBA from ThemeColor');
		assert.strictEqual(themeColor.GetHex(), '#c0504d', 'Check Hex from ThemeColor');
		
		assert.equalRgb(autoColor.GetRGB(), {r : 0, g : 0, b : 0}, 'AutoColor is black in RGB');
		assert.equalRgba(autoColor.GetRGBA(), {r : 0, g : 0, b : 0, a : 255}, 'AutoColor is black in RGBA');
		assert.strictEqual(autoColor.GetHex(), '#000000', 'AutoColor is black in Hex');
	});
});
