/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
function rgb(r, g, b, a) {
	a = a || 0;
	return {R: r, G:g, B: b, A: a};
}
function test(startColor, arrMods, expectedColor, description) {
	const oMods = new AscFormat.CColorModifiers();
	for (let i = 0; i < arrMods.length; i += 1) {
		const mod = arrMods[i];
		oMods.addMod(mod.name, mod.value);
	}
	oMods.Apply(startColor);
	return {expected: expectedColor, result: startColor, description: description};
}
function mod(name, value) {
	return {name: name, value: value};
}
	QUnit.module("Test applying color mods");
	const tests = [
		test(
			rgb(68, 114, 196),
			[mod('satOff', 0), mod('lumOff', 0), mod('hueOff')],
			rgb(68, 114, 196),
			"Check uselessness of zero offsets"
		),
		test(
			rgb(68, 114, 196),
			[],
			rgb(68, 114, 196),
			"Check color with empty mods"
		),
		test(
			rgb(68, 114, 196),
			[mod("hueMod", 100000), mod("satMod", 100000), mod("lumMod", 100000)],
			rgb(68, 114, 196),
			"Check uselessness of multiplication by 100%"
		),
		test(
			rgb(68, 114, 196),
			[mod("satMod", 5000)],
			rgb(129, 131, 135),
			"Check saturation multiplication by 5%"
		),
		test(
			rgb(68, 114, 196),
			[mod("satMod", 10000)],
			rgb(126, 130, 138),
			"Check saturation multiplication by 10%"
		),
		test(
			rgb(68, 114, 196),
			[mod("lumMod", 10000)],
			rgb(6, 11, 20),
			"Check lightness multiplication by 10%"
		),
		test(
			rgb(68, 114, 196),
			[mod("lumMod", 20000)],
			rgb(13, 23, 40),
			"Check lightness multiplication by 20%"
		),
	];


	QUnit.test('Check colors', (assert) => {
		for (let i = 0; i < tests.length; i++) {
			const test = tests[i];
			assert.deepEqual(test.result, test.expected, test.description);
		}
	});

});
