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
function rgb(r, g, b, a) {
	a = typeof a === 'number' ? a : 255;
	return {R: r, G:g, B: b, A: a};
}
function test(startColor, arrMods, expectedColor) {
	const oMods = new AscFormat.CColorModifiers();
	const description = "Check  applying ";
	const modsDescription = [];
	for (let i = 0; i < arrMods.length; i += 1) {
		const mod = arrMods[i];
		oMods.addMod(mod.name, mod.value);
		modsDescription.push(mod.description);
	}
	oMods.Apply(startColor);
	return {expected: expectedColor, result: startColor, description: description + modsDescription.join(', ')};
}
function mod(name, value) {
	return {name: name, value: value, description: name + " with " + value + " value"};
}
	QUnit.module("Test applying color mods");
	const tests = [
		test(
			rgb(68, 114, 196),
			[mod('satOff', 0), mod('lumOff', 0), mod('hueOff', 0)],
			rgb(68, 114, 196)
		),
		test(
			rgb(68, 114, 196),
			[],
			rgb(68, 114, 196)
		),
		test(
			rgb(68, 114, 196),
			[mod("hueMod", 100000), mod("satMod", 100000), mod("lumMod", 100000)],
			rgb(68, 114, 196)
		),
		test(
			rgb(68, 114, 196),
			[mod("satMod", 5000)],
			rgb(129, 131, 135)
		),
		test(
			rgb(68, 114, 196),
			[mod("satMod", 10000)],
			rgb(126, 130, 138)
		),
		test(
			rgb(68, 114, 196),
			[mod("lumMod", 10000)],
			rgb(6, 11, 20)
		),
		test(
			rgb(68, 114, 196),
			[mod("lumMod", 20000)],
			rgb(13, 23, 40)
		),
		test(
			rgb(68, 114, 196),
			[mod("lumOff", 20000)],
			rgb(146, 172, 220)
		),
		test(
			rgb(68, 114, 196),
			[mod("hueOff", 20000)],
			rgb(68, 113, 196)
		),
		test(
			rgb(68, 114, 196),
			[mod("hueOff", -200000)],
			rgb(68, 121, 196)
		),
		test(
			rgb(165, 165, 165),
			[mod("hueOff", 677650)],
			rgb(165, 165, 165)
		),

		test(
			rgb(165, 165, 165),
			[mod("satOff", 0), mod("lumOff", 0), mod("hueOff", 0), mod("alphaOff", 0)],
			rgb(165, 165, 165)
		),

		test(
			rgb(134, 86, 64),
			[mod("hueOff", 2710599)],
			rgb(129, 134, 64)
		),
		test(
			rgb(165, 165, 165),
			[mod("lumOff", -3676)],
			rgb(156, 156, 156)
		),
		test(
			rgb(98, 168, 87),
			[mod("tint", 12000)],
			rgb(243, 247, 242)
		),
		test(
			rgb(98, 168, 87),
			[mod("tint", 50000)],
			rgb(197, 217, 195)
		),
		test(
			rgb(157, 54, 14),
			[mod("hueMod", 44000)],
			rgb(157, 32, 14)
		)
	];

	const todoTests = [
		test(
			rgb(157, 54, 14),
			[mod("satMod", 200000)],
			rgb(229, 22, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 0)],
			rgb(85, 85, 85)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 10000)],
			rgb(93, 82, 78)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 20000)],
			rgb(100, 79, 71)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 30000)],
			rgb(107, 76, 64)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 100000)],
			rgb(157, 54, 14)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 150000)],
			rgb(193, 38, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 110000)],
			rgb(164, 51, 7)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 120000)],
			rgb(171, 48, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 130000)],
			rgb(178, 45, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 130000)],
			rgb(178, 45, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 140000)],
			rgb(186, 41, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 150000)],
			rgb(193, 38, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 160000)],
			rgb(200, 35, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 170000)],
			rgb(207, 32, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 180000)],
			rgb(214, 29, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 190000)],
			rgb(221, 26, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 200000)],
			rgb(229, 22, 0)
		),
		test(
			rgb(157, 54, 14),
			[mod("satOff", 10000)],
			rgb(166, 50, 5)
		),
		test(
			rgb(157, 54, 14),
			[mod("satOff", 20000)],
			rgb(174, 46, 0)
		),
		test(
			rgb(165, 165, 165),
			[mod("satOff", 25000)],
			rgb(187, 142, 53)
		),
		test(
			rgb(165, 165, 165),
			[mod("satOff", 40000)],
			rgb(201, 129, 0)
		),
		test(
			rgb(165, 165, 165),
			[mod("satOff", 10000)],
			rgb(174, 156, 120)
		),
		test(
			rgb(165, 165, 165),
			[mod("satOff", 30000)],
			rgb(192, 138, 30)
		),
		test(
			rgb(165, 165, 165),
			[mod("satOff", 35000)],
			rgb(196, 134, 8)
		),
		test(
			rgb(165, 165, 165),
			[mod("hueOff", 677650), mod("satOff", 25000), mod("lumOff", -3676),  mod("alphaOff", 0)],
			rgb(173, 131, 48)
		),
		test(
			rgb(165, 165, 165),
			[mod("hueOff", 1355300), mod("satOff", 50000), mod("lumOff", -7353),  mod("alphaOff", 0)],
			rgb(173, 99, 0)
		),
		test(
			rgb(165, 165, 165),
			[mod("hueOff", 2032949), mod("satOff", 75000), mod("lumOff", -11029),  mod("alphaOff", 0)],
			rgb(176, 74, 0)
		),
		test(
			rgb(165, 165, 165),
			[mod("hueOff", 2710599), mod("satOff", 100000), mod("lumOff", -14706),  mod("alphaOff", 0)],
			rgb(180, 53, 0)
		),
		test(
			rgb(131, 131, 131),
			[mod("satOff", 99200)],
			rgb(254, 8, 0)
		),
		test(
			rgb(127, 127, 127),
			[mod("satOff", 99200)],
			rgb(253, 1, 0)
		),
		test(
			rgb(223, 219, 213),
			[mod("shade", 80000)],
			rgb(202, 198, 193)
		),
		test(
			rgb(223, 219, 213),
			[mod("shade", 50000)],
			rgb(70, 122, 62)
		),
		test(
			rgb(98, 168, 87),
			[mod("tint", 90000)],
			rgb(126, 179, 119)
		),
	];
	QUnit.test('Check colors with mods', (assert) => {
		for (let i = 0; i < tests.length; i++) {
			const test = tests[i];
			assert.deepEqual(test.result, test.expected, test.description);
		}
	});
});
