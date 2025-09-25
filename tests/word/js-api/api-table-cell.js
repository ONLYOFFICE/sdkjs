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
	QUnit.module('Test the ApiTableCell methods');

	QUnit.test('SetColor, GetColor', function (assert) {
		const table = AscTest.Editor.CreateTable(2, 2);
		const cell = table.GetCell(0, 0);

		assert.strictEqual(cell.GetBackgroundColor(), null, 'Color check for a newly created table cell');

		cell.SetBackgroundColor(255, 127, 0);
		assert.equalRgb(cell.GetBackgroundColor(), { r: 255, g: 127, b: 0 }, 'Color check after setting color with RGB components');

		const hexColor = AscTest.Editor.HexColor('bada55');
		cell.SetBackgroundColor(hexColor);
		assert.strictEqual(cell.GetBackgroundColor().GetHex(), '#bada55', 'Color check after setting color with ApiColor (hex)');

		const themeColor = AscTest.Editor.ThemeColor('accent2');
		cell.SetBackgroundColor(themeColor);
		assert.strictEqual(cell.GetBackgroundColor().IsThemeColor(), true, 'Color check after setting color with ApiColor (theme)');

		cell.SetBackgroundColor(0, 0, 0, true);
		assert.strictEqual(cell.GetBackgroundColor(), null, 'Color check after resetting color');
	});
});
