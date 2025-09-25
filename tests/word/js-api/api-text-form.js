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
	QUnit.module('Test the ApiTextForm methods');

	QUnit.test('SetBorderColor, GetBorderColor', function (assert) {
		const textForm = AscTest.Editor.CreateTextForm();

		assert.strictEqual(textForm.GetBorderColor(), null, 'Check border color for a newly created text form');

		textForm.SetBorderColor(255, 122, 100);
		assert.equalRgb(textForm.GetBorderColor(), { r: 255, g: 122, b: 100 }, 'Check border color after setting it with rgba components');

		const hexColor = AscTest.Editor.HexColor('a1b2c3');
		textForm.SetBorderColor(hexColor);
		assert.equalRgb(textForm.GetBorderColor(), { r: 161, g: 178, b: 195 }, 'Check border color after setting it with ApiColor (rgba)');

		textForm.SetBorderColor(0, 0, 0, true);
		assert.strictEqual(textForm.GetBorderColor(), null, 'Check border color after resetting it');
	});

	QUnit.test('SetBackgroundColor, GetBackgroundColor', function (assert) {
		const textForm = AscTest.Editor.CreateTextForm();

		assert.strictEqual(textForm.GetBackgroundColor(), null, 'Check background color for a newly created text form');

		textForm.SetBackgroundColor(255, 122, 100);
		assert.equalRgb(textForm.GetBackgroundColor(), { r: 255, g: 122, b: 100 }, 'Check background color after setting it with rgba components');

		const hexColor = AscTest.Editor.HexColor('a1b2c3');
		textForm.SetBackgroundColor(hexColor);
		assert.equalRgb(textForm.GetBackgroundColor().GetRGB(), { r: 161, g: 178, b: 195 }, 'Check background color after setting it with ApiColor (rgba)');

		const themeColor = AscTest.Editor.ThemeColor('accent3');
		textForm.SetBackgroundColor(themeColor);
		assert.strictEqual(textForm.GetBackgroundColor().IsThemeColor(), true, 'Check background color after setting it with theme color');

		textForm.SetBackgroundColor(0, 0, 0, true);
		assert.strictEqual(textForm.GetBackgroundColor(), null, 'Check background color after resetting it');
	});
});
