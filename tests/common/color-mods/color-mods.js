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

const IS_APPROX_EQUALS = false;
$(function () {
function rgb(r, g, b, a) {
	a = typeof a === 'number' ? a : 255;
	return {R: r, G:g, B: b, A: a};
}
	const assertTest = (assert) => {
		return (testObject) => {
			if (IS_APPROX_EQUALS) {
				assert.true(AscFormat.fApproxEqual(testObject.result.R, testObject.expected.R, 1.001), testObject.description + ": R");
				assert.true(AscFormat.fApproxEqual(testObject.result.G, testObject.expected.G, 1.001), testObject.description + ": G");
				assert.true(AscFormat.fApproxEqual(testObject.result.B, testObject.expected.B, 1.001), testObject.description + ": B");
			} else {
				assert.deepEqual(testObject.result, testObject.expected, testObject.description);
			}

		}
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
			rgb(157, 54, 14),
			[mod("hueMod", 44000)],
			rgb(157, 32, 14)
		),
		test(
			rgb(157, 54, 14),
			[mod("satMod", 200000)],
			rgb(229, 22, 0)
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
			[mod("hueOff", 2710599), mod("satOff", 100000), mod("lumOff", -14706),  mod("alphaOff", 0)],
			rgb(180, 53, 0)
		),
		test(
			rgb(131, 131, 131),
			[mod("satOff", 99200)],
			rgb(254, 8, 0)
		),

		test(rgb(127, 127, 127),[mod("satOff", 100000)],rgb(254, 0, 0)),
		test(rgb(0, 0, 0),[mod("satOff", 100000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satOff", 100000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satOff", 100000)],rgb(30, 0, 0)),
		test(rgb(127, 127, 127),[mod("satOff", 100000)],rgb(254, 0, 0)),
		test(rgb(126, 126, 126),[mod("satOff", 100000)],rgb(252, 0, 0)),
		test(rgb(128, 128, 128),[mod("satOff", 100000)],rgb(255, 1, 0)),
		test(rgb(200, 200, 200),[mod("satOff", 100000)],rgb(255, 145, 0)),
		test(rgb(127, 127, 127),[mod("satOff", 150000)],rgb(255, 0, 0)),
		test(rgb(0, 0, 0),[mod("satOff", 150000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satOff", 150000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satOff", 150000)],rgb(38, 0, 0)),
		test(rgb(127, 127, 127),[mod("satOff", 150000)],rgb(255, 0, 0)),
		test(rgb(126, 126, 126),[mod("satOff", 150000)],rgb(255, 0, 0)),
		test(rgb(128, 128, 128),[mod("satOff", 150000)],rgb(255, 0, 0)),
		test(rgb(200, 200, 200),[mod("satOff", 150000)],rgb(255, 117, 0)),
		test(rgb(127, 127, 127),[mod("satOff", 5000)],rgb(133, 121, 95)),
		test(rgb(0, 0, 0),[mod("satOff", 5000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satOff", 5000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satOff", 5000)],rgb(16, 14, 11)),
		test(rgb(127, 127, 127),[mod("satOff", 5000)],rgb(133, 121, 95)),
		test(rgb(126, 126, 126),[mod("satOff", 5000)],rgb(132, 120, 95)),
		test(rgb(128, 128, 128),[mod("satOff", 5000)],rgb(134, 122, 96)),
		test(rgb(200, 200, 200),[mod("satOff", 5000)],rgb(203, 197, 186)),
		test(rgb(127, 127, 127),[mod("satOff", -77000)],rgb(29, 225, 255)),
		test(rgb(0, 0, 0),[mod("satOff", -77000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satOff", -77000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satOff", -77000)],rgb(3, 27, 73)),
		test(rgb(127, 127, 127),[mod("satOff", -77000)],rgb(29, 225, 255)),
		test(rgb(126, 126, 126),[mod("satOff", -77000)],rgb(29, 223, 255)),
		test(rgb(128, 128, 128),[mod("satOff", -77000)],rgb(30, 226, 255)),
		test(rgb(200, 200, 200),[mod("satOff", -77000)],rgb(158, 242, 255)),
		test(rgb(127, 127, 127),[mod("satOff", -5000)],rgb(121, 133, 159)),
		test(rgb(0, 0, 0),[mod("satOff", -5000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satOff", -5000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satOff", -5000)],rgb(14, 16, 19)),
		test(rgb(127, 127, 127),[mod("satOff", -5000)],rgb(121, 133, 159)),
		test(rgb(126, 126, 126),[mod("satOff", -5000)],rgb(120, 132, 158)),
		test(rgb(128, 128, 128),[mod("satOff", -5000)],rgb(122, 134, 160)),
		test(rgb(200, 200, 200),[mod("satOff", -5000)],rgb(197, 203, 214)),

		test(rgb(127, 127, 127),[mod("satMod", 50000)],rgb(127, 127, 127)),
		test(rgb(0, 0, 0),[mod("satMod", 50000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satMod", 50000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satMod", 50000)],rgb(15, 15, 15)),
		test(rgb(127, 127, 127),[mod("satMod", 50000)],rgb(127, 127, 127)),
		test(rgb(126, 126, 126),[mod("satMod", 50000)],rgb(126, 126, 126)),
		test(rgb(128, 128, 128),[mod("satMod", 50000)],rgb(128, 128, 128)),
		test(rgb(200, 200, 200),[mod("satMod", 50000)],rgb(200, 200, 200)),
		test(rgb(127, 127, 127),[mod("satMod", 150000)],rgb(127, 127, 127)),
		test(rgb(0, 0, 0),[mod("satMod", 150000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satMod", 150000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satMod", 150000)],rgb(15, 15, 15)),
		test(rgb(127, 127, 127),[mod("satMod", 150000)],rgb(127, 127, 127)),
		test(rgb(126, 126, 126),[mod("satMod", 150000)],rgb(126, 126, 126)),
		test(rgb(128, 128, 128),[mod("satMod", 150000)],rgb(128, 128, 128)),
		test(rgb(200, 200, 200),[mod("satMod", 150000)],rgb(200, 200, 200)),
		test(rgb(127, 127, 127),[mod("satMod", 5000)],rgb(127, 127, 127)),
		test(rgb(0, 0, 0),[mod("satMod", 5000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satMod", 5000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satMod", 5000)],rgb(15, 15, 15)),
		test(rgb(127, 127, 127),[mod("satMod", 5000)],rgb(127, 127, 127)),
		test(rgb(126, 126, 126),[mod("satMod", 5000)],rgb(126, 126, 126)),
		test(rgb(128, 128, 128),[mod("satMod", 5000)],rgb(128, 128, 128)),
		test(rgb(200, 200, 200),[mod("satMod", 5000)],rgb(200, 200, 200)),
		test(rgb(127, 127, 127),[mod("satMod", -77000)],rgb(127, 127, 127)),
		test(rgb(0, 0, 0),[mod("satMod", -77000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satMod", -77000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satMod", -77000)],rgb(15, 15, 15)),
		test(rgb(127, 127, 127),[mod("satMod", -77000)],rgb(127, 127, 127)),
		test(rgb(126, 126, 126),[mod("satMod", -77000)],rgb(126, 126, 126)),
		test(rgb(128, 128, 128),[mod("satMod", -77000)],rgb(128, 128, 128)),
		test(rgb(200, 200, 200),[mod("satMod", -77000)],rgb(200, 200, 200)),
		test(rgb(127, 127, 127),[mod("satMod", -5000)],rgb(127, 127, 127)),
		test(rgb(0, 0, 0),[mod("satMod", -5000)],rgb(0, 0, 0)),
		test(rgb(255, 255, 255),[mod("satMod", -5000)],rgb(255, 255, 255)),
		test(rgb(15, 15, 15),[mod("satMod", -5000)],rgb(15, 15, 15)),
		test(rgb(127, 127, 127),[mod("satMod", -5000)],rgb(127, 127, 127)),
		test(rgb(126, 126, 126),[mod("satMod", -5000)],rgb(126, 126, 126)),
		test(rgb(128, 128, 128),[mod("satMod", -5000)],rgb(128, 128, 128)),
		test(rgb(200, 200, 200),[mod("satMod", -5000)],rgb(200, 200, 200)),
	];

	const todoTests = [
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
			[mod("satOff", 25000)],
			rgb(187, 142, 53)
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
			[mod("satMod", 0)],
			rgb(85, 85, 85)
		),
	];
	QUnit.test('Check colors with mods', (assert) => {
		const fTestFunction = assertTest(assert);
		for (let i = 0; i < tests.length; i++) {
			const test = tests[i];
			fTestFunction(test);
		}
	});
	QUnit.test.todo('Check colors with mods', (assert) => {
		const fTestFunction = assertTest(assert);
		for (let i = 0; i < todoTests.length; i++) {
			const test = todoTests[i];
			fTestFunction(test);
		}
	});


	QUnit.test('Check colors with satMod', (assert) => {
		//test mods
		// [mod("satMod", 0)]
		// [mod("satMod", 1)]
		// [mod("satMod", 2)]
		// [mod("satMod", 10)]
		// [mod("satMod", 100)]
		// [mod("satMod", 1000)]
		// [mod("satMod", 5000)]
		// [mod("satMod", 10000)]
		// [mod("satMod", 50000)]
		// [mod("satMod", 100000)]
		// [mod("satMod", 150000)]
		// [mod("satMod", 200000)]
		// [mod("satMod", 1000000)]
		// [mod("satMod", -1000000)]
		// [mod("satMod", -500000)]
		// [mod("satMod", -150000)]
		// [mod("satMod", -100000)]
		// [mod("satMod", -50000)]
		// [mod("satMod", -10000)]
		// [mod("satMod", -5000)]
		// [mod("satMod", -1000)]
		// [mod("satMod", -500)]
		// [mod("satMod", -100)]
		// [mod("satMod", -50)]
		// [mod("satMod", -10)]
		// [mod("satMod", -2)]
		// [mod("satMod", -1)]
		//test colors
		// rgb(0,0,0)
		const fTestFunction = assertTest(assert);
		let testResult;
		testResult = test(rgb(0,0,0),[mod("satMod", 0)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 1)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 2)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 10)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 100)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 1000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 5000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 10000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 50000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 100000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 150000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 200000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", 1000000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -1000000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -500000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -150000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -100000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -50000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -10000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -5000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -1000)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -500)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -100)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -50)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -10)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -2)],rgb(0, 0, 0));	
		fTestFunction(testResult);
		testResult = test(rgb(0,0,0),[mod("satMod", -1)],rgb(0, 0, 0));	
		fTestFunction(testResult);
			//.assert/deepEqual(test(rgb(126,126,126),)
		testResult = test(rgb(126,126,126),[mod("satMod", 0)],rgb(126, 126, 126));	
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 1)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 2)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 10)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 100)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 1000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 5000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 10000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 50000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 100000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 150000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 200000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", 1000000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -1000000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -500000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -150000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -100000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -50000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -10000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -5000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -1000)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -500)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -100)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -50)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -10)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -2)],rgb(126, 126, 126));
		fTestFunction(testResult);
		testResult = test(rgb(126,126,126),[mod("satMod", -1)],rgb(126, 126, 126));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(127,127,127),)
		testResult = test(rgb(127,127,127),[mod("satMod", 0)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 1)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 2)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 10)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 100)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 1000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 5000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 10000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 50000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 100000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 150000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 200000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", 1000000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -1000000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -500000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -150000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -100000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -50000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -10000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -5000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -1000)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -500)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -100)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -50)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -10)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -2)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(127,127,127),[mod("satMod", -1)],rgb(127, 127, 127));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(128,128,128),)
		testResult = test(rgb(128,128,128),[mod("satMod", 0)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 1)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 2)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 10)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 100)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 1000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 5000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 10000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 50000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 100000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 150000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 200000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", 1000000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -1000000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -500000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -150000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -100000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -50000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -10000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -5000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -1000)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -500)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -100)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -50)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -10)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -2)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(128,128,128),[mod("satMod", -1)],rgb(128, 128, 128));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(200,200,200),)
		testResult = test(rgb(200,200,200),[mod("satMod", 0)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 1)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 2)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 10)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 100)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 1000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 5000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 10000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 50000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 100000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 150000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 200000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", 1000000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -1000000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -500000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -150000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -100000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -50000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -10000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -5000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -1000)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -500)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -100)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -50)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -10)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -2)],rgb(200, 200, 200));
		fTestFunction(testResult);
		testResult = test(rgb(200,200,200),[mod("satMod", -1)],rgb(200, 200, 200));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(255,255,255),)
		testResult = test(rgb(255,255,255),[mod("satMod", 0)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 1)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 2)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 10)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 100)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 1000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 5000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 10000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 50000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 100000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 150000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 200000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", 1000000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -1000000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -500000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -150000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -100000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -50000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -10000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -5000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -1000)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -500)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -100)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -50)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -10)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -2)],rgb(255, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255,255,255),[mod("satMod", -1)],rgb(255, 255, 255));
		fTestFunction(testResult);
		//.assert/deepEqual(test(1gb(101,100,100),)
		testResult = test(rgb(100,100,100),[mod("satMod", 0)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 1)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 2)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 10)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 100)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 1000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 5000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 10000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 50000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 100000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 150000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 200000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", 1000000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -1000000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -500000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -150000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -100000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -50000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -10000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -5000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -1000)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -500)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -100)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -50)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -10)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -2)],rgb(100, 100, 100));
		fTestFunction(testResult);
		testResult = test(rgb(100,100,100),[mod("satMod", -1)],rgb(100, 100, 100));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(34, 139, 34),)
		testResult = test(rgb(34, 139, 34),[mod("satMod", 0)],rgb(87, 87, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 1)],rgb(87, 87, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 2)],rgb(87, 87, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 10)],rgb(86, 87, 86));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 100)],rgb(86, 87, 86));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 1000)],rgb(86, 87, 86));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 5000)],rgb(84, 89, 84));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 10000)],rgb(81, 92, 81));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 50000)],rgb(60, 113, 60));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 100000)],rgb(34, 139, 34));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 150000)],rgb(8, 165, 8));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 200000)],rgb(0, 191, 0));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", 1000000)],rgb(0, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -1000000)],rgb(255, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -500000)],rgb(255, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -150000)],rgb(165, 8, 165));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -100000)],rgb(139, 34, 139));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -50000)],rgb(113, 60, 113));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -10000)],rgb(92, 81, 92));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -5000)],rgb(89, 84, 89));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -1000)],rgb(87, 86, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -500)],rgb(87, 86, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -100)],rgb(87, 86, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -50)],rgb(87, 86, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -10)],rgb(87, 86, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -2)],rgb(87, 87, 87));
		fTestFunction(testResult);
		testResult = test(rgb(34, 139, 34),[mod("satMod", -1)],rgb(87, 87, 87));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(255, 99, 71),)
		testResult = test(rgb(255, 99, 71),[mod("satMod", 0)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 1)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 2)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 10)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 100)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 1000)],rgb(164, 162, 162));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 5000)],rgb(168, 160, 158));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 10000)],rgb(172, 157, 154));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 50000)],rgb(209, 131, 117));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 100000)],rgb(255, 99, 71));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 150000)],rgb(255, 67, 25));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 200000)],rgb(255, 35, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", 1000000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -1000000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -500000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -150000)],rgb(25, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -100000)],rgb(71, 227, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -50000)],rgb(117, 195, 209));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -10000)],rgb(154, 169, 172));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -5000)],rgb(158, 166, 168));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -1000)],rgb(162, 164, 164));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -500)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -100)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -50)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -10)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -2)],rgb(163, 163, 163));
		fTestFunction(testResult);
		testResult = test(rgb(255, 99, 71),[mod("satMod", -1)],rgb(163, 163, 163));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(75, 0, 130),)
		testResult = test(rgb(75, 0, 130),[mod("satMod", 0)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 1)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 2)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 10)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 100)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 1000)],rgb(65, 64, 66));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 5000)],rgb(66, 62, 68));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 10000)],rgb(66, 59, 71));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 50000)],rgb(70, 32, 98));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 100000)],rgb(75, 0, 130));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 150000)],rgb(80, 0, 163));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 200000)],rgb(85, 0, 195));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", 1000000)],rgb(165, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -1000000)],rgb(0, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -500000)],rgb(15, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -150000)],rgb(50, 163, 0));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -100000)],rgb(55, 130, 0));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -50000)],rgb(60, 98, 32));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -10000)],rgb(64, 71, 59));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -5000)],rgb(64, 68, 62));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -1000)],rgb(65, 66, 64));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -500)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -100)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -50)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -10)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -2)],rgb(65, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(75, 0, 130),[mod("satMod", -1)],rgb(65, 65, 65));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(220, 20, 60),)
		testResult = test(rgb(220, 20, 60),[mod("satMod", 0)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 1)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 2)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 10)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 100)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 1000)],rgb(121, 119, 119));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 5000)],rgb(125, 115, 117));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 10000)],rgb(130, 110, 114));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 50000)],rgb(170, 70, 90));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 100000)],rgb(220, 20, 60));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 150000)],rgb(255, 0, 30));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 200000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", 1000000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -1000000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -500000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -150000)],rgb(0, 255, 210));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -100000)],rgb(20, 220, 180));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -50000)],rgb(70, 170, 150));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -10000)],rgb(110, 130, 126));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -5000)],rgb(115, 125, 123));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -1000)],rgb(119, 121, 121));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -500)],rgb(120, 121, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -100)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -50)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -10)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -2)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(220, 20, 60),[mod("satMod", -1)],rgb(120, 120, 120));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(0, 191, 255),)
		testResult = test(rgb(0, 191, 255),[mod("satMod", 0)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 1)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 2)],rgb(127, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 10)],rgb(127, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 100)],rgb(127, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 1000)],rgb(126, 128, 129));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 5000)],rgb(121, 131, 134));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 10000)],rgb(115, 134, 140));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 50000)],rgb(64, 159, 191));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 100000)],rgb(0, 191, 255));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 150000)],rgb(0, 223, 255));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 200000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", 1000000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -1000000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -500000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -150000)],rgb(255, 32, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -100000)],rgb(255, 64, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -50000)],rgb(191, 96, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -10000)],rgb(140, 121, 115));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -5000)],rgb(134, 124, 121));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -1000)],rgb(129, 127, 126));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -500)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -100)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -50)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -10)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -2)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(0, 191, 255),[mod("satMod", -1)],rgb(128, 127, 127));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(255, 215, 0),)
		testResult = test(rgb(255, 215, 0),[mod("satMod", 0)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 1)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 2)],rgb(128, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 10)],rgb(128, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 100)],rgb(128, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 1000)],rgb(129, 128, 126));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 5000)],rgb(134, 132, 121));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 10000)],rgb(140, 136, 115));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 50000)],rgb(191, 171, 64));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 100000)],rgb(255, 215, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 150000)],rgb(255, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 200000)],rgb(255, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", 1000000)],rgb(255, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -1000000)],rgb(0, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -500000)],rgb(0, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -150000)],rgb(0, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -100000)],rgb(0, 40, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -50000)],rgb(64, 84, 191));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -10000)],rgb(115, 119, 140));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -5000)],rgb(121, 123, 134));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -1000)],rgb(126, 127, 129));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -500)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -100)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -50)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -10)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -2)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 215, 0),[mod("satMod", -1)],rgb(127, 127, 128));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(138, 43, 226),)
		testResult = test(rgb(138, 43, 226),[mod("satMod", 0)],rgb(135, 135, 135));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 1)],rgb(135, 134, 135));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 2)],rgb(135, 134, 135));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 10)],rgb(135, 134, 135));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 100)],rgb(135, 134, 135));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 1000)],rgb(135, 134, 135));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 5000)],rgb(135, 130, 139));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 10000)],rgb(135, 125, 144));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 50000)],rgb(136, 89, 180));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 100000)],rgb(138, 43, 226));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 150000)],rgb(140, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 200000)],rgb(141, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", 1000000)],rgb(169, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -1000000)],rgb(100, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -500000)],rgb(117, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -150000)],rgb(129, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -100000)],rgb(131, 226, 43));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -50000)],rgb(133, 180, 89));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -10000)],rgb(134, 144, 125));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -5000)],rgb(134, 139, 130));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -1000)],rgb(134, 135, 134));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -500)],rgb(134, 135, 134));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -100)],rgb(134, 135, 134));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -50)],rgb(134, 135, 134));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -10)],rgb(135, 135, 134));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -2)],rgb(135, 135, 134));
		fTestFunction(testResult);
		testResult = test(rgb(138, 43, 226),[mod("satMod", -1)],rgb(135, 135, 134));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(50, 205, 50),)
		testResult = test(rgb(50, 205, 50),[mod("satMod", 0)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 1)],rgb(128, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 2)],rgb(127, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 10)],rgb(127, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 100)],rgb(127, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 1000)],rgb(127, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 5000)],rgb(124, 131, 124));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 10000)],rgb(120, 135, 120));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 50000)],rgb(89, 166, 89));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 100000)],rgb(50, 205, 50));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 150000)],rgb(11, 244, 11));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 200000)],rgb(0, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", 1000000)],rgb(0, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -1000000)],rgb(255, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -500000)],rgb(255, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -150000)],rgb(244, 11, 244));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -100000)],rgb(205, 50, 205));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -50000)],rgb(166, 89, 166));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -10000)],rgb(135, 120, 135));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -5000)],rgb(131, 124, 131));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -1000)],rgb(128, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -500)],rgb(128, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -100)],rgb(128, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -50)],rgb(128, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -10)],rgb(128, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -2)],rgb(128, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(50, 205, 50),[mod("satMod", -1)],rgb(128, 128, 128));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(255, 69, 0),)
		testResult = test(rgb(255, 69, 0),[mod("satMod", 0)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 1)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 2)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 10)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 100)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 1000)],rgb(129, 127, 126));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 5000)],rgb(134, 125, 121));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 10000)],rgb(140, 122, 115));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 50000)],rgb(191, 98, 64));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 100000)],rgb(255, 69, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 150000)],rgb(255, 40, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 200000)],rgb(255, 11, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", 1000000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -1000000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -500000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -150000)],rgb(0, 215, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -100000)],rgb(0, 186, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -50000)],rgb(64, 157, 191));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -10000)],rgb(115, 133, 140));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -5000)],rgb(121, 130, 134));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -1000)],rgb(126, 128, 129));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -500)],rgb(127, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -100)],rgb(127, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -50)],rgb(127, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -10)],rgb(127, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -2)],rgb(127, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 69, 0),[mod("satMod", -1)],rgb(127, 127, 128));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(0, 128, 128),)
		testResult = test(rgb(0, 128, 128),[mod("satMod", 0)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 1)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 2)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 10)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 100)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 1000)],rgb(63, 65, 65));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 5000)],rgb(61, 67, 67));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 10000)],rgb(58, 70, 70));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 50000)],rgb(32, 96, 96));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 100000)],rgb(0, 128, 128));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 150000)],rgb(0, 160, 160));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 200000)],rgb(0, 192, 192));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", 1000000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -1000000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -500000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -150000)],rgb(160, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -100000)],rgb(128, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -50000)],rgb(96, 32, 32));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -10000)],rgb(70, 58, 58));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -5000)],rgb(67, 61, 61));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -1000)],rgb(65, 63, 63));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -500)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -100)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -50)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -10)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -2)],rgb(64, 64, 64));
		fTestFunction(testResult);
		testResult = test(rgb(0, 128, 128),[mod("satMod", -1)],rgb(64, 64, 64));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(218, 112, 214),)
		testResult = test(rgb(218, 112, 214),[mod("satMod", 0)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 1)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 2)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 10)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 100)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 1000)],rgb(166, 164, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 5000)],rgb(168, 162, 167));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 10000)],rgb(170, 160, 170));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 50000)],rgb(191, 138, 189));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 100000)],rgb(218, 112, 214));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 150000)],rgb(244, 86, 238));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 200000)],rgb(255, 59, 255));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", 1000000)],rgb(255, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -1000000)],rgb(0, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -500000)],rgb(0, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -150000)],rgb(86, 244, 92));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -100000)],rgb(112, 218, 116));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -50000)],rgb(138, 191, 141));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -10000)],rgb(160, 170, 160));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -5000)],rgb(162, 168, 163));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -1000)],rgb(164, 166, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -500)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -100)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -50)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -10)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -2)],rgb(165, 165, 165));
		fTestFunction(testResult);
		testResult = test(rgb(218, 112, 214),[mod("satMod", -1)],rgb(165, 165, 165));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(70, 130, 180),)
		testResult = test(rgb(70, 130, 180),[mod("satMod", 0)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 1)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 2)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 10)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 100)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 1000)],rgb(124, 125, 126));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 5000)],rgb(122, 125, 128));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 10000)],rgb(120, 126, 131));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 50000)],rgb(98, 127, 152));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 100000)],rgb(70, 130, 180));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 150000)],rgb(43, 132, 207));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 200000)],rgb(15, 135, 235));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", 1000000)],rgb(0, 175, 255));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -1000000)],rgb(255, 75, 0));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -500000)],rgb(255, 100, 0));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -150000)],rgb(207, 118, 43));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -100000)],rgb(180, 120, 70));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -50000)],rgb(152, 123, 98));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -10000)],rgb(131, 125, 120));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -5000)],rgb(128, 125, 122));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -1000)],rgb(126, 125, 124));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -500)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -100)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -50)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -10)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -2)],rgb(125, 125, 125));
		fTestFunction(testResult);
		testResult = test(rgb(70, 130, 180),[mod("satMod", -1)],rgb(125, 125, 125));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(255, 165, 0),)
		testResult = test(rgb(255, 165, 0),[mod("satMod", 0)],rgb(127, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 1)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 2)],rgb(128, 127, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 10)],rgb(128, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 100)],rgb(128, 128, 127));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 1000)],rgb(129, 128, 126));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 5000)],rgb(134, 129, 121));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 10000)],rgb(140, 131, 115));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 50000)],rgb(191, 146, 64));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 100000)],rgb(255, 165, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 150000)],rgb(255, 184, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 200000)],rgb(255, 202, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", 1000000)],rgb(255, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -1000000)],rgb(0, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -500000)],rgb(0, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -150000)],rgb(0, 71, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -100000)],rgb(0, 90, 255));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -50000)],rgb(64, 109, 191));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -10000)],rgb(115, 124, 140));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -5000)],rgb(121, 126, 134));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -1000)],rgb(126, 127, 129));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -500)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -100)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -50)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -10)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -2)],rgb(127, 127, 128));
		fTestFunction(testResult);
		testResult = test(rgb(255, 165, 0),[mod("satMod", -1)],rgb(127, 127, 128));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(147, 112, 219),)
		testResult = test(rgb(147, 112, 219),[mod("satMod", 0)],rgb(166, 166, 166));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 1)],rgb(166, 166, 166));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 2)],rgb(166, 166, 166));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 10)],rgb(165, 165, 166));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 100)],rgb(165, 165, 166));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 1000)],rgb(165, 165, 166));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 5000)],rgb(165, 163, 168));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 10000)],rgb(164, 160, 171));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 50000)],rgb(156, 139, 192));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 100000)],rgb(147, 112, 219));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 150000)],rgb(138, 85, 246));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 200000)],rgb(129, 59, 255));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", 1000000)],rgb(0, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -1000000)],rgb(255, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -500000)],rgb(255, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -150000)],rgb(193, 246, 85));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -100000)],rgb(184, 219, 112));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -50000)],rgb(175, 192, 139));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -10000)],rgb(167, 171, 160));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -5000)],rgb(166, 168, 163));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -1000)],rgb(166, 166, 165));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -500)],rgb(166, 166, 165));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -100)],rgb(166, 166, 165));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -50)],rgb(166, 166, 165));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -10)],rgb(166, 166, 165));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -2)],rgb(166, 166, 166));
		fTestFunction(testResult);
		testResult = test(rgb(147, 112, 219),[mod("satMod", -1)],rgb(166, 166, 166));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(60, 179, 113),)
		testResult = test(rgb(60, 179, 113),[mod("satMod", 0)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 1)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 2)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 10)],rgb(119, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 100)],rgb(119, 120, 119));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 1000)],rgb(119, 120, 119));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 5000)],rgb(117, 122, 119));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 10000)],rgb(114, 125, 119));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 50000)],rgb(90, 149, 116));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 100000)],rgb(60, 179, 113));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 150000)],rgb(30, 209, 110));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 200000)],rgb(0, 239, 106));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", 1000000)],rgb(0, 255, 54));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -1000000)],rgb(255, 0, 185));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -500000)],rgb(255, 0, 152));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -150000)],rgb(209, 30, 129));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -100000)],rgb(179, 60, 126));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -50000)],rgb(149, 90, 123));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -10000)],rgb(125, 114, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -5000)],rgb(122, 117, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -1000)],rgb(120, 119, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -500)],rgb(120, 119, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -100)],rgb(120, 119, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -50)],rgb(120, 119, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -10)],rgb(120, 119, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -2)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(60, 179, 113),[mod("satMod", -1)],rgb(120, 120, 120));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(210, 105, 30),)
		testResult = test(rgb(210, 105, 30),[mod("satMod", 0)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 1)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 2)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 10)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 100)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 1000)],rgb(121, 120, 119));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 5000)],rgb(124, 119, 115));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 10000)],rgb(129, 118, 111));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 50000)],rgb(165, 112, 75));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 100000)],rgb(210, 105, 30));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 150000)],rgb(255, 97, 0));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 200000)],rgb(255, 90, 0));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", 1000000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -1000000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -500000)],rgb(0, 195, 255));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -150000)],rgb(0, 142, 255));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -100000)],rgb(30, 135, 210));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -50000)],rgb(75, 127, 165));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -10000)],rgb(111, 121, 129));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -5000)],rgb(115, 121, 124));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -1000)],rgb(119, 120, 121));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -500)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -100)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -50)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -10)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -2)],rgb(120, 120, 120));
		fTestFunction(testResult);
		testResult = test(rgb(210, 105, 30),[mod("satMod", -1)],rgb(120, 120, 120));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(123, 104, 238),)
		testResult = test(rgb(123, 104, 238),[mod("satMod", 0)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 1)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 2)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 10)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 100)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 1000)],rgb(171, 170, 172));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 5000)],rgb(169, 168, 174));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 10000)],rgb(166, 164, 178));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 50000)],rgb(147, 137, 204));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 100000)],rgb(123, 104, 238));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 150000)],rgb(99, 70, 255));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 200000)],rgb(75, 37, 255));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", 1000000)],rgb(0, 0, 255));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -1000000)],rgb(255, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -500000)],rgb(255, 255, 0));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -150000)],rgb(243, 255, 70));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -100000)],rgb(219, 238, 104));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -50000)],rgb(195, 204, 137));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -10000)],rgb(176, 178, 164));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -5000)],rgb(173, 174, 168));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -1000)],rgb(171, 172, 170));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -500)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -100)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -50)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -10)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -2)],rgb(171, 171, 171));
		fTestFunction(testResult);
		testResult = test(rgb(123, 104, 238),[mod("satMod", -1)],rgb(171, 171, 171));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(0, 206, 209),)
		testResult = test(rgb(0, 206, 209),[mod("satMod", 0)],rgb(104, 104, 104));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 1)],rgb(104, 105, 105));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 2)],rgb(104, 105, 105));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 10)],rgb(104, 105, 105));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 100)],rgb(104, 105, 105));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 1000)],rgb(103, 106, 106));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 5000)],rgb(99, 110, 110));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 10000)],rgb(94, 115, 115));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 50000)],rgb(52, 155, 157));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 100000)],rgb(0, 206, 209));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 150000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 200000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", 1000000)],rgb(0, 255, 255));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -1000000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -500000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -150000)],rgb(255, 0, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -100000)],rgb(209, 3, 0));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -50000)],rgb(157, 54, 52));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -10000)],rgb(115, 94, 94));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -5000)],rgb(110, 99, 99));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -1000)],rgb(106, 103, 103));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -500)],rgb(105, 104, 104));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -100)],rgb(105, 104, 104));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -50)],rgb(105, 104, 104));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -10)],rgb(105, 104, 104));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -2)],rgb(105, 104, 104));
		fTestFunction(testResult);
		testResult = test(rgb(0, 206, 209),[mod("satMod", -1)],rgb(105, 104, 104));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(255, 105, 180),)
		testResult = test(rgb(255, 105, 180),[mod("satMod", 0)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 1)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 2)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 10)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 100)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 1000)],rgb(181, 179, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 5000)],rgb(184, 176, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 10000)],rgb(187, 173, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 50000)],rgb(217, 142, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 100000)],rgb(255, 105, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 150000)],rgb(255, 67, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 200000)],rgb(255, 30, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", 1000000)],rgb(255, 0, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -1000000)],rgb(0, 255, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -500000)],rgb(0, 255, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -150000)],rgb(67, 255, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -100000)],rgb(105, 255, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -50000)],rgb(142, 217, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -10000)],rgb(173, 187, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -5000)],rgb(176, 184, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -1000)],rgb(179, 181, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -500)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -100)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -50)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -10)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -2)],rgb(180, 180, 180));
		fTestFunction(testResult);
		testResult = test(rgb(255, 105, 180),[mod("satMod", -1)],rgb(180, 180, 180));
		fTestFunction(testResult);
		//.assert/deepEqual(test(rgb(46, 139, 87),)
		testResult = test(rgb(46, 139, 87),[mod("satMod", 0)],rgb(93, 93, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 1)],rgb(93, 93, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 2)],rgb(92, 93, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 10)],rgb(92, 93, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 100)],rgb(92, 93, 92));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 1000)],rgb(92, 93, 92));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 5000)],rgb(90, 95, 92));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 10000)],rgb(88, 97, 92));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 50000)],rgb(69, 116, 90));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 100000)],rgb(46, 139, 87));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 150000)],rgb(23, 162, 84));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 200000)],rgb(0, 185, 82));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", 1000000)],rgb(0, 255, 38));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -1000000)],rgb(255, 0, 147));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -500000)],rgb(255, 0, 120));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -150000)],rgb(162, 23, 101));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -100000)],rgb(139, 46, 98));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -50000)],rgb(116, 69, 95));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -10000)],rgb(97, 88, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -5000)],rgb(95, 90, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -1000)],rgb(93, 92, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -500)],rgb(93, 92, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -100)],rgb(93, 92, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -50)],rgb(93, 92, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -10)],rgb(93, 92, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -2)],rgb(93, 92, 93));
		fTestFunction(testResult);
		testResult = test(rgb(46, 139, 87),[mod("satMod", -1)],rgb(93, 93, 93));
		fTestFunction(testResult);
	});
});
